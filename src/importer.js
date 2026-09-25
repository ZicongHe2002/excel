import { unzipSync } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

const MAX_FILE_BYTES = 80 * 1024 * 1024;
const MAX_EXPANDED_BYTES = 250 * 1024 * 1024;
const MAX_XML_BYTES = 32 * 1024 * 1024;
const SKU_HEADERS = new Set(['款号', '货号', '商品货号', '商品款号', '款式编码', '商品编码', 'SKU']);
let nextWorkbookId = 1;

function children(node, name) {
  return Array.from(node?.childNodes || []).filter((item) => item.nodeType === 1 && (!name || item.localName === name));
}

function descendants(node, name) {
  return Array.from(node.getElementsByTagName('*')).filter((item) => item.localName === name);
}

function textOf(node) {
  return node?.textContent || '';
}

function parseXml(bytes, path) {
  if (!bytes) throw new Error(`Excel 缺少必要文件：${path}`);
  if (bytes.length > MAX_XML_BYTES) throw new Error('表格内容过大，请拆分后重新上传。');
  const utf16 = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le' : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  const source = new TextDecoder(utf16).decode(bytes);
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error('此文件的 XML 结构不受支持，请用 Excel 或 WPS 另存为 .xlsx 后重试。');
  let invalid = false;
  const doc = new DOMParser({ errorHandler: { warning() {}, error() { invalid = true; }, fatalError() { invalid = true; } } }).parseFromString(source, 'application/xml');
  if (invalid || !doc.documentElement) throw new Error(`Excel 内容损坏，无法读取 ${path}。`);
  return doc;
}

function resolvePath(base, target) {
  if (!target || /^(?:[a-z]+:|\/\/)/i.test(target)) return null;
  let decoded;
  try { decoded = decodeURIComponent(target); } catch { decoded = target; }
  const pieces = decoded.startsWith('/') ? [] : base.split('/').slice(0, -1);
  for (const part of decoded.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!pieces.length) return null; pieces.pop(); }
    else pieces.push(part);
  }
  return pieces.join('/');
}

function relsPath(path) {
  const slash = path.lastIndexOf('/');
  return `${slash < 0 ? '' : path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
}

function relationships(files, part) {
  const path = part ? relsPath(part) : '_rels/.rels';
  if (!files[path]) return [];
  return descendants(parseXml(files[path], path), 'Relationship').map((rel) => ({
    id: rel.getAttribute('Id'),
    type: rel.getAttribute('Type').split('/').pop(),
    path: rel.getAttribute('TargetMode') === 'External' ? null : resolvePath(part, rel.getAttribute('Target')),
  }));
}

// Check central-directory sizes before fflate allocates decompressed buffers.
function inspectZip(bytes) {
  if (bytes.length > MAX_FILE_BYTES) throw new Error('单个文件不能超过 80 MB，请拆分后上传。');
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('暂不支持旧版 .xls 或加密文件，请取消密码并另存为 .xlsx。');
  if (bytes.length < 22 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('这不是有效的 .xlsx 文件，请用 Excel 或 WPS 另存为 .xlsx。');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let position = bytes.length - 22; position >= Math.max(0, bytes.length - 65557); position--) {
    if (view.getUint32(position, true) === 0x06054b50 && position + 22 + view.getUint16(position + 20, true) === bytes.length) { end = position; break; }
  }
  if (end < 0) throw new Error('Excel 文件不完整或已损坏，请重新保存后上传。');
  const entries = view.getUint16(end + 10, true);
  const offset = view.getUint32(end + 16, true);
  if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) || offset === 0xffffffff || entries === 0xffff) throw new Error('暂不支持分卷或超大型 Excel，请拆分为普通 .xlsx 文件。');
  if (entries > 10000) throw new Error('文件包含过多内部资源，请拆分后上传。');
  let cursor = offset;
  let expanded = 0;
  const paths = new Set();
  for (let index = 0; index < entries; index++) {
    if (cursor + 46 > end || view.getUint32(cursor, true) !== 0x02014b50) throw new Error('Excel 文件目录损坏，请重新保存后上传。');
    if (view.getUint16(cursor + 8, true) & 1) throw new Error('请先取消 Excel 文件密码，再重新上传。');
    const size = view.getUint32(cursor + 24, true);
    const filenameLength = view.getUint16(cursor + 28, true);
    const entryEnd = cursor + 46 + filenameLength + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
    if (entryEnd > end || size === 0xffffffff) throw new Error('Excel 文件目录损坏或文件过大。');
    const path = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + filenameLength));
    if (path.startsWith('/') || path.includes('\\') || path.split('/').some((segment) => segment === '..') || paths.has(path)) throw new Error('Excel 内部路径异常，请重新另存为 .xlsx 后上传。');
    paths.add(path);
    expanded += size;
    if (size > MAX_FILE_BYTES || expanded > MAX_EXPANDED_BYTES || (/\.(xml|rels)$/i.test(path) && size > MAX_XML_BYTES)) throw new Error('Excel 解压后内容过大，请拆分文件后上传。');
    cursor = entryEnd;
  }
}

function richText(node) {
  // Phonetic guides (rPh) are metadata, not part of the displayed identifier.
  return children(node).map((item) => item.localName === 't' ? textOf(item) : item.localName === 'r' ? children(item, 't').map(textOf).join('') : '').join('');
}

function readNumberFormats(files, path) {
  if (!path || !files[path]) return [];
  const doc = parseXml(files[path], path);
  const formats = new Map(descendants(doc, 'numFmt').map((item) => [item.getAttribute('numFmtId'), item.getAttribute('formatCode')]));
  const xfs = children(doc.documentElement, 'cellXfs')[0];
  return children(xfs, 'xf').map((item) => formats.get(item.getAttribute('numFmtId')) || '');
}

function cellText(cell, strings, numberFormats) {
  const type = cell.getAttribute('t');
  const raw = textOf(children(cell, 'v')[0]);
  if (type === 's') return strings[Number(raw)] || '';
  if (type === 'inlineStr') return richText(children(cell, 'is')[0]);
  if (type === 'b') return raw === '1' ? 'TRUE' : 'FALSE';
  if (!type || type === 'n') {
    const format = numberFormats[Number(cell.getAttribute('s') || 0)];
    // Preserve common all-digit SKUs displayed as 000012; avoid guessing complex formats.
    if (format && /^0+$/.test(format) && /^\d+$/.test(raw)) return raw.padStart(format.length, '0');
  }
  return raw;
}

function columnNumber(column) {
  return [...column].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function parseRange(value) {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(value || '');
  if (!match) return null;
  return { startColumn: columnNumber(match[1].toUpperCase()), start: Number(match[2]), endColumn: columnNumber((match[3] || match[1]).toUpperCase()), end: Number(match[4] || match[2]) };
}

function headerKey(value) {
  return value.trim().toUpperCase().replace(/[\s：:]/g, '');
}

function sourceSku(value) {
  const lines = value.trim().split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  if (lines.length === 2 && /^(?:男|女|男款|女款)$/.test(lines[1])) return lines[0];
  return value.trim();
}

function readSheet(files, info, strings, numberFormats) {
  const doc = parseXml(files[info.path], info.path);
  const sheetData = children(doc.documentElement, 'sheetData')[0];
  const rows = children(sheetData, 'row').map((row, index) => ({
    number: Number(row.getAttribute('r') || index + 1),
    cells: children(row, 'c').map((cell) => ({ cell, column: /^([A-Z]+)/i.exec(cell.getAttribute('r'))?.[1]?.toUpperCase(), value: cellText(cell, strings, numberFormats) })),
  })).sort((a, b) => a.number - b.number);
  const headers = [];
  for (const row of rows.filter((item) => item.number <= 30)) {
    for (const cell of row.cells) {
      if (!cell.column || !SKU_HEADERS.has(headerKey(cell.value))) continue;
      const score = row.cells.filter((item) => /^(序号|图片|材质|颜色|尺码|价格|库存|商品名称)/.test(item.value.trim())).length;
      headers.push({ row: row.number, column: cell.column, score });
    }
  }
  headers.sort((a, b) => b.score - a.score || a.row - b.row);
  if (!headers.length) return { warning: `“${info.name}”未找到“款号 / 货号 / SKU”表头，已跳过。` };
  const header = headers[0];
  const skuColumnNumber = columnNumber(header.column);
  const merges = descendants(doc, 'mergeCell').map((item) => parseRange(item.getAttribute('ref'))).filter(Boolean);
  const skuMerges = merges.filter((range) => range.startColumn === skuColumnNumber && range.endColumn === skuColumnNumber).sort((a, b) => a.start - b.start);
  const mergeByStart = new Map(skuMerges.map((range) => [range.start, range]));
  const headerMerge = mergeByStart.get(header.row);
  const headerEnd = headerMerge?.end || header.row;
  const warnings = [];
  const starts = [];
  let lastContentRow = headerEnd;
  let mergeIndex = 0;
  for (const row of rows) {
    while (mergeIndex < skuMerges.length && skuMerges[mergeIndex].end < row.number) mergeIndex++;
    if (row.cells.some((item) => item.value.trim() || children(item.cell, 'f').length)) lastContentRow = Math.max(lastContentRow, row.number);
    if (row.number <= headerEnd) continue;
    const sku = row.cells.find((item) => item.column === header.column)?.value.trim();
    if (!sku) continue;
    if (SKU_HEADERS.has(headerKey(sku))) {
      starts.push({ repeatedHeader: true, start: row.number });
      continue;
    }
    if (/^(?:合计|总计|小计|备注|说明)(?:[：:].*)?$/.test(sku)) {
      starts.push({ footer: true, start: row.number });
      continue;
    }
    const candidateMerge = skuMerges[mergeIndex];
    const containingMerge = candidateMerge && candidateMerge.start < row.number && candidateMerge.end >= row.number;
    if (containingMerge) continue;
    starts.push({ sku, normalizedSku: normalizeSku(sourceSku(sku)), start: row.number });
  }
  const blocks = [];
  for (let index = 0; index < starts.length; index++) {
    const item = starts[index];
    if (item.repeatedHeader || item.footer) continue;
    const nextStart = starts[index + 1]?.start;
    const merged = mergeByStart.get(item.start);
    const end = merged ? merged.end : nextStart ? nextStart - 1 : lastContentRow;
    if (nextStart && end >= nextStart) throw new Error(`“${info.name}”第 ${item.start} 行的款号合并区域与下一款重叠，请先修正原表。`);
    blocks.push({ ...item, end: Math.max(item.start, end) });
  }
  if (!blocks.length) return { warning: `“${info.name}”找到表头，但没有可提取的货号。` };
  if (starts.some((item) => item.repeatedHeader)) warnings.push(`“${info.name}”含重复表头，导出仅保留顶部表头。`);
  return { name: info.name, path: info.path, doc, skuColumn: header.column, headerEnd, blocks, warnings };
}

export function normalizeSku(value) {
  return String(value ?? '').trim().toUpperCase();
}

export function parseSkuInput(value) {
  const codes = [];
  const duplicates = [];
  const seen = new Set();
  const seenDuplicates = new Set();
  for (const token of String(value ?? '').split(/[\s,，;；、]+/).filter(Boolean)) {
    const code = normalizeSku(token);
    if (seen.has(code)) {
      if (!seenDuplicates.has(code)) { duplicates.push(code); seenDuplicates.add(code); }
    } else { codes.push(code); seen.add(code); }
  }
  return { codes, duplicates };
}

export function readWorkbook(input, { name = '商品资料.xlsx', id } = {}) {
  if (/\.xls$/i.test(name)) throw new Error('暂不支持旧版 .xls，请用 Excel 或 WPS 另存为 .xlsx 后上传。');
  if (!/\.xlsx$/i.test(name)) throw new Error('请上传 .xlsx 格式的 Excel 文件。');
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  inspectZip(bytes);
  let files;
  try { files = unzipSync(bytes); } catch { throw new Error('Excel 文件解压失败，可能已损坏或受到密码保护。'); }
  if (!files['[Content_Types].xml']) throw new Error('文件不是有效的 Excel 工作簿。');
  const packageRels = relationships(files, '');
  const workbookPath = packageRels.find((rel) => rel.type === 'officeDocument')?.path;
  if (!workbookPath || !files[workbookPath]) throw new Error('文件缺少 Excel 工作簿内容，请重新另存为 .xlsx。');
  const workbookDoc = parseXml(files[workbookPath], workbookPath);
  if (workbookDoc.documentElement.localName !== 'workbook') throw new Error('此文件不是 Excel 工作簿。');
  const workbookRels = relationships(files, workbookPath);
  const sharedStringsPath = workbookRels.find((rel) => rel.type === 'sharedStrings')?.path;
  const stylesPath = workbookRels.find((rel) => rel.type === 'styles')?.path || null;
  const themePath = workbookRels.find((rel) => rel.type === 'theme')?.path || null;
  const sharedStrings = sharedStringsPath && files[sharedStringsPath] ? descendants(parseXml(files[sharedStringsPath], sharedStringsPath), 'si') : [];
  const strings = sharedStrings.map(richText);
  const numberFormats = readNumberFormats(files, stylesPath);
  const sheets = [];
  const warnings = [];
  for (const sheet of descendants(workbookDoc, 'sheet')) {
    const relationshipId = sheet.getAttribute('r:id') || Array.from(sheet.attributes).find((attribute) => attribute.localName === 'id')?.value;
    const rel = workbookRels.find((item) => item.id === relationshipId);
    if (!rel?.path || rel.type !== 'worksheet') { warnings.push(`“${sheet.getAttribute('name')}”不是可读取的数据工作表，已跳过。`); continue; }
    const result = readSheet(files, { name: sheet.getAttribute('name'), path: rel.path }, strings, numberFormats);
    if (result.warning) warnings.push(result.warning);
    else { sheets.push(result); warnings.push(...result.warnings); }
  }
  if (!sheets.length) throw new Error('没有找到含商品数据的工作表。请确认前 30 行内有“款号”“货号”或“SKU”表头。');
  return {
    id: id ?? `workbook-${nextWorkbookId++}`,
    name, files, sheets, sharedStrings, sharedStringsPath, stylesPath, themePath, workbookPath, warnings,
    skuCount: new Set(sheets.flatMap((sheet) => sheet.blocks.map((block) => block.normalizedSku))).size,
    rowCount: sheets.reduce((total, sheet) => total + sheet.blocks.reduce((sum, block) => sum + block.end - block.start + 1, 0), 0),
  };
}

export function matchSelection(workbooks, codes) {
  const requested = [...new Set(codes.map(normalizeSku).filter(Boolean))];
  const wanted = new Set(requested);
  const found = new Set();
  const selections = [];
  let blockCount = 0;
  let rowCount = 0;
  for (const workbook of workbooks) {
    for (const sheet of workbook.sheets) {
      const blocks = sheet.blocks.filter((block) => wanted.has(block.normalizedSku));
      if (!blocks.length) continue;
      selections.push({ workbook, sheet, blocks });
      for (const block of blocks) { found.add(block.normalizedSku); blockCount++; rowCount += block.end - block.start + 1; }
    }
  }
  return { selections, matchedCodes: requested.filter((code) => found.has(code)), missingCodes: requested.filter((code) => !found.has(code)), blockCount, rowCount };
}
