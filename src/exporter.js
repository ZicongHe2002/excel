import { zipSync } from 'fflate';
import { matchSelection, parseSkuInput } from './importer.js';
import { mergeStyles } from './styles.js';
import { createCellImageCopier, CELL_IMAGE_REL, CELL_IMAGE_TYPE } from './cell-images.js';
import { NS, parseXml, serialize, xmlBytes, children, all, first, remove, el, readRels, relationships, relsPath, cellRef, rangeRef, mapRange, columnNumber, columnName } from './xml.js';

const CT = 'application/vnd.openxmlformats-officedocument.';
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp' };
const sheetMetrics = new WeakMap();

function imageCatalog(workbook) {
  const catalog = new Map();
  for (const path of Object.keys(workbook.files).filter(p => /(^|\/)cellimages\.xml$/i.test(p))) {
    const doc = parseXml(workbook.files[path]);
    const rels = readRels(workbook.files, path);
    for (const pic of all(doc, 'pic')) {
      const name = all(pic, 'cNvPr')[0]?.getAttribute('name');
      const blip = all(pic, 'blip')[0];
      const rel = rels.find(r => r.id === (blip?.getAttributeNS(NS.rel, 'embed') || blip?.getAttribute('r:embed')));
      if (name && rel && !rel.external) catalog.set(name, { pic, path: rel.path });
    }
  }
  return catalog;
}

function imageSize(bytes, pic) {
  if (bytes?.length > 24 && bytes[0] === 137 && bytes[1] === 80) {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return [view.getUint32(16), view.getUint32(20)];
  }
  if (bytes?.[0] === 0xff && bytes[1] === 0xd8) {
    let pos = 2;
    while (pos + 9 < bytes.length) {
      if (bytes[pos++] !== 0xff) continue;
      while (bytes[pos] === 0xff) pos++;
      const marker = bytes[pos++];
      if (marker === 0xd9 || marker === 0xda) break;
      if (marker === 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
      const size = bytes[pos] * 256 + bytes[pos + 1];
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) return [bytes[pos + 5] * 256 + bytes[pos + 6], bytes[pos + 3] * 256 + bytes[pos + 4]];
      if (size < 2) break;
      pos += size;
    }
  }
  const extent = all(pic, 'xfrm').flatMap(n => children(n, 'ext'))[0];
  return [Number(extent?.getAttribute('cx') || 1), Number(extent?.getAttribute('cy') || 1)];
}

function colPixels(root, number) {
  const col = all(root, 'col').find(n => Number(n.getAttribute('min')) <= number && Number(n.getAttribute('max')) >= number);
  if (col?.getAttribute('hidden') === '1') return 0;
  const width = Number(col?.getAttribute('width') || first(root, 'sheetFormatPr')?.getAttribute('defaultColWidth') || 8.43);
  const digitWidth = sheetMetrics.get(root) || 7;
  // ISO 29500 column widths already contain cell padding; do not add it twice.
  return Math.floor((256 * width + Math.floor(128 / digitWidth)) / 256 * digitWidth);
}
function rowPixels(root, row) {
  const node = children(first(root, 'sheetData'), 'row').find(n => Number(n.getAttribute('r')) === row);
  if (node?.getAttribute('hidden') === '1') return 0;
  return Number(node?.getAttribute('ht') || first(root, 'sheetFormatPr')?.getAttribute('defaultRowHeight') || 15) * 96 / 72;
}

function positionFromAbsolute(anchor, root) {
  const pos = first(anchor, 'pos');
  let x = Number(pos?.getAttribute('x') || 0), y = Number(pos?.getAttribute('y') || 0);
  let row = 1, col = 1;
  const heights = new Map(children(first(root, 'sheetData'), 'row').map(n => [Number(n.getAttribute('r')), n.getAttribute('hidden') === '1' ? 0 : Number(n.getAttribute('ht') || first(root, 'sheetFormatPr')?.getAttribute('defaultRowHeight') || 15) * 12700]));
  const defaultHeight = Number(first(root, 'sheetFormatPr')?.getAttribute('defaultRowHeight') || 15) * 12700;
  while (row < 1048576) { const h = heights.get(row) ?? defaultHeight; if (h > 0 && y < h) break; y -= h; row++; }
  while (col < 16384) { const w = colPixels(root, col) * 9525; if (w > 0 && x < w) break; x -= w; col++; }
  return { row, col, x: Math.max(0, x), y: Math.max(0, y) };
}

function synchronizePictureGeometry(anchor, root, originalAnchor = null, originalRoot = null) {
  const from = first(anchor, 'from');
  if (!from) return;
  function point(marker, grid = root) {
    let x = Number(first(marker, 'colOff')?.textContent || 0);
    let y = Number(first(marker, 'rowOff')?.textContent || 0);
    for (let col = 1; col <= Number(first(marker, 'col')?.textContent || 0); col++) x += colPixels(grid, col) * 9525;
    for (let row = 1; row <= Number(first(marker, 'row')?.textContent || 0); row++) y += rowPixels(grid, row) * 9525;
    return { x: Math.round(x), y: Math.round(y) };
  }
  const start = point(from);
  const to = first(anchor, 'to');
  const end = to ? point(to) : null;
  const extent = first(anchor, 'ext');
  const width = end ? end.x - start.x : Number(extent?.getAttribute('cx'));
  const height = end ? end.y - start.y : Number(extent?.getAttribute('cy'));
  for (const pic of all(anchor, 'pic')) {
    const transform = all(pic, 'xfrm')[0];
    if (!transform) continue;
    const off = first(transform, 'off');
    if (originalAnchor && originalRoot && off) {
      const oldFrom = first(originalAnchor, 'from');
      const oldTo = first(originalAnchor, 'to');
      const oldY = oldFrom ? point(oldFrom, originalRoot).y : Number(first(originalAnchor, 'pos')?.getAttribute('y') || 0);
      // Source columns are unchanged. Preserve native drawing X/width instead
      // of replacing accurate source geometry with font-metric estimates.
      off.setAttribute('y', Number(off.getAttribute('y')) + start.y - oldY);
      if (oldTo && end) {
        const size = first(transform, 'ext');
        size?.setAttribute('cy', Number(size.getAttribute('cy')) + height - (point(oldTo, originalRoot).y - oldY));
      }
      continue;
    }
    off?.setAttribute('x', start.x); off?.setAttribute('y', start.y);
    const size = first(transform, 'ext');
    if (width > 0) size?.setAttribute('cx', Math.round(width));
    if (height > 0) size?.setAttribute('cy', Math.round(height));
  }
}

function makeCellAnchor(drawing, source, bytes, ref, root, relId, pictureId) {
  const pos = cellRef(ref);
  const merge = all(root, 'mergeCell').map(n => rangeRef(n.getAttribute('ref'))).find(r => r.start.row === pos.row && r.start.column === pos.column);
  const lastRow = merge?.end.row || pos.row;
  const lastCol = merge?.end.column || pos.column;
  let width = 0, height = 0;
  for (let col = pos.column; col <= lastCol; col++) width += colPixels(root, col);
  for (let row = pos.row; row <= lastRow; row++) height += rowPixels(root, row);
  const [iw, ih] = imageSize(bytes, source.pic);
  const scale = Math.min(Math.max(1, width - 4) / iw, Math.max(1, height - 4) / ih);
  const w = iw * scale, h = ih * scale;
  const cx = Math.round(w * 9525), cy = Math.round(h * 9525);
  const anchor = el(drawing, 'xdr:oneCellAnchor', {}, NS.drawing);
  const from = anchor.appendChild(el(drawing, 'xdr:from', {}, NS.drawing));
  for (const [name, value] of [['col', pos.column - 1], ['colOff', Math.round((width - w) / 2 * 9525)], ['row', pos.row - 1], ['rowOff', Math.round((height - h) / 2 * 9525)]]) {
    const node = from.appendChild(el(drawing, `xdr:${name}`, {}, NS.drawing)); node.appendChild(drawing.createTextNode(String(value)));
  }
  anchor.appendChild(el(drawing, 'xdr:ext', { cx, cy }, NS.drawing));
  const pic = drawing.importNode(source.pic, true);
  const info = all(pic, 'cNvPr')[0]; info?.setAttribute('id', pictureId);
  for (const blip of all(pic, 'blip')) blip.setAttributeNS(NS.rel, 'r:embed', relId);
  for (const transform of all(pic, 'xfrm')) {
    const off = first(transform, 'off'); off?.setAttribute('x', 0); off?.setAttribute('y', 0);
    const ext = first(transform, 'ext'); ext?.setAttribute('cx', cx); ext?.setAttribute('cy', cy);
  }
  anchor.appendChild(pic);
  anchor.appendChild(el(drawing, 'xdr:clientData', {}, NS.drawing));
  synchronizePictureGeometry(anchor, root);
  return anchor;
}

function uniqueSheetName(name, used) {
  const base = name.replace(/\.xlsx$/i, '').replace(/[\\/*?:\[\]]/g, '_').replace(/^'+|'+$/g, '').trim() || '选款';
  let n = 1, candidate = base.slice(0, 31);
  while (used.has(candidate.toLowerCase())) { const suffix = ` (${++n})`; candidate = base.slice(0, 31 - suffix.length) + suffix; }
  used.add(candidate.toLowerCase()); return candidate;
}

function moveFormulaToValue(cell, warnings) {
  const formula = first(cell, 'f');
  if (!formula) return;
  const value = first(cell, 'v');
  if (!value || value.textContent === '') throw new Error(`单元格 ${cell.getAttribute('r')} 的公式没有已保存的结果，请先在 Excel / WPS 中计算并保存原文件。`);
  remove(formula);
  warnings.add('普通公式已保留原文件中保存的计算结果，导出表可独立使用。');
}

export async function exportSelection(workbooks, inputCodes, { imageMode = 'original' } = {}) {
  if (!['original', 'drawing'].includes(imageMode)) throw new Error('不支持的图片导出方式。');
  const codes = typeof inputCodes === 'string' ? parseSkuInput(inputCodes).codes : inputCodes;
  const result = matchSelection(workbooks, codes || []);
  if (!result.selections.length) throw new Error('没有找到对应货号，请检查货号或上传的商品表。');
  const sources = [...new Set(result.selections.map(s => s.workbook))];
  const styles = mergeStyles(sources);
  const output = { 'xl/styles.xml': xmlBytes(styles.doc) };
  const warnings = new Set();
  const catalogs = imageMode === 'drawing' ? new Map(sources.map(w => [w, imageCatalog(w)])) : null;
  const mediaMaps = new Map(sources.map(w => [w, new Map()]));
  const typeOverrides = [{ path: '/xl/workbook.xml', type: `${CT}spreadsheetml.sheet.main+xml` }, { path: '/xl/styles.xml', type: `${CT}spreadsheetml.styles+xml` }];
  const mediaTypes = new Map();
  const workbookDoc = parseXml(`<workbook xmlns="${NS.sheet}" xmlns:r="${NS.rel}"><bookViews><workbookView/></bookViews><sheets/></workbook>`);
  const wbSheets = first(workbookDoc.documentElement, 'sheets');
  const wbRels = [{ id: 'rStyles', type: `${NS.rel}/styles`, target: 'styles.xml' }];
  const usedNames = new Set();
  let mediaCounter = 0, imageCount = 0;
  function copyMedia(workbook, path) {
    if (mediaMaps.get(workbook).has(path)) return mediaMaps.get(workbook).get(path);
    if (!workbook.files[path]) throw new Error(`原表缺少图片文件：${path}，请重新保存原表后再试。`);
    const ext = path.split('.').at(-1).toLowerCase();
    if (!MIME[ext]) throw new Error(`暂不支持 ${ext} 格式的图片，请先在原表中转换为 PNG 或 JPG。`);
    const target = `xl/media/image${++mediaCounter}.${ext}`;
    output[target] = workbook.files[path]; mediaTypes.set(ext, MIME[ext]);
    mediaMaps.get(workbook).set(path, target);
    return target;
  }
  const nativeImages = imageMode === 'original' ? createCellImageCopier(sources, copyMedia) : null;
  for (const [sheetIndex, selection] of result.selections.entries()) {
    await new Promise(resolve => setTimeout(resolve, 0));
    const { workbook, sheet, blocks } = selection;
    const number = sheetIndex + 1;
    const styleMap = styles.maps.get(workbook);
    const rowNumbers = new Set(Array.from({ length: sheet.headerEnd }, (_, i) => i + 1));
    for (const block of blocks) for (let r = block.start; r <= block.end; r++) rowNumbers.add(r);
    const rowMap = new Map([...rowNumbers].sort((a, b) => a - b).map((r, i) => [r, i + 1]));
    const doc = parseXml(serialize(sheet.doc));
    const root = doc.documentElement;
    sheetMetrics.set(root, styleMap.maximumDigitWidth);
    sheetMetrics.set(sheet.doc.documentElement, styleMap.maximumDigitWidth);
    root.setAttribute('xmlns:r', NS.rel);
    const sheetData = first(root, 'sheetData');
    const sourceRels = readRels(workbook.files, sheet.path);
    const drawingDoc = parseXml(`<xdr:wsDr xmlns:xdr="${NS.drawing}" xmlns:a="${NS.art}" xmlns:r="${NS.rel}"/>`);
    const drawingRels = [];
    const sheetRels = [];
    const pendingImages = [];
    const sourceCols = all(root, 'col').map(n => ({ min: +n.getAttribute('min'), max: +n.getAttribute('max'), style: +(n.getAttribute('style') || 0) }));
    for (const row of children(sheetData, 'row')) {
      const oldRow = Number(row.getAttribute('r'));
      if (!rowMap.has(oldRow)) { remove(row); continue; }
      row.setAttribute('r', rowMap.get(oldRow));
      const originalRowStyle = Number(row.getAttribute('s') || 0);
      if (row.hasAttribute('s')) row.setAttribute('s', originalRowStyle + styleMap.cellXfs);
      for (const cell of children(row, 'c')) {
        const ref = cellRef(cell.getAttribute('r'));
        if (!ref) continue;
        const originalStyle = cell.hasAttribute('s') ? +cell.getAttribute('s') : (row.getAttribute('customFormat') === '1' ? originalRowStyle : sourceCols.find(c => c.min <= ref.column && c.max >= ref.column)?.style || 0);
        cell.setAttribute('r', `${ref.col}${rowMap.get(oldRow)}`);
        cell.setAttribute('s', originalStyle + styleMap.cellXfs);
        cell.removeAttribute('cm'); cell.removeAttribute('vm');
        if (cell.getAttribute('t') === 's') {
          const index = Number(first(cell, 'v')?.textContent);
          const shared = workbook.sharedStrings[index];
          if (!shared) throw new Error('原表的文字引用不完整，请重新保存原表后再试。');
          remove(first(cell, 'v'));
          const inline = el(doc, 'is');
          for (const node of Array.from(shared.childNodes)) inline.appendChild(doc.importNode(node, true));
          cell.appendChild(inline); cell.setAttribute('t', 'inlineStr');
        }
        const formula = first(cell, 'f');
        const imgMatch = /(?:_xlfn\.)?DISPIMG\(\s*"([^"]+)"/i.exec(formula?.textContent || first(cell, 'v')?.textContent || '');
        if (imgMatch) {
          if (nativeImages) {
            nativeImages.copyCell(workbook, cell, imgMatch[1]);
            imageCount++;
            continue;
          }
          const image = catalogs.get(workbook).get(imgMatch[1]);
          if (!image) throw new Error(`${workbook.name} 的 ${cell.getAttribute('r')} 缺少对应原图，请重新保存包含图片的 Excel。`);
          pendingImages.push({ ref: cell.getAttribute('r'), image });
          remove(formula); remove(first(cell, 'v')); cell.removeAttribute('t');
        } else moveFormulaToValue(cell, warnings);
      }
    }
    for (const col of all(root, 'col')) if (col.hasAttribute('style')) col.setAttribute('style', Number(col.getAttribute('style')) + styleMap.cellXfs);
    const merges = first(root, 'mergeCells');
    for (const merge of children(merges)) {
      const mapped = mapRange(merge.getAttribute('ref'), rowMap);
      if (!mapped || !mapped.includes(':')) remove(merge); else merge.setAttribute('ref', mapped);
    }
    if (merges) { merges.setAttribute('count', children(merges).length); if (!children(merges).length) remove(merges); }
    for (const conditional of children(root, 'conditionalFormatting')) {
      const refs = conditional.getAttribute('sqref').split(/\s+/).map(r => mapRange(r, rowMap)).filter(Boolean);
      if (!refs.length) { remove(conditional); continue; }
      conditional.setAttribute('sqref', refs.join(' '));
      for (const rule of all(conditional, 'cfRule')) {
        if (rule.hasAttribute('dxfId')) rule.setAttribute('dxfId', Number(rule.getAttribute('dxfId')) + styleMap.dxfs);
        if (all(rule, 'formula').length) { remove(conditional); warnings.add('原表的公式条件格式未迁移，普通单元格格式已保留。'); break; }
      }
    }
    for (const link of all(root, 'hyperlink')) {
      const mapped = mapRange(link.getAttribute('ref'), rowMap);
      if (!mapped) { remove(link); continue; }
      link.setAttribute('ref', mapped);
      const rid = link.getAttributeNS(NS.rel, 'id') || link.getAttribute('r:id');
      const relation = sourceRels.find(r => r.id === rid);
      if (relation?.external) {
        const id = `rLink${sheetRels.length + 1}`;
        link.setAttributeNS(NS.rel, 'r:id', id); sheetRels.push({ ...relation, id });
      } else if (rid) remove(link);
    }
    for (const drawingNode of children(root, 'drawing')) {
      const relation = sourceRels.find(r => r.id === (drawingNode.getAttributeNS(NS.rel, 'id') || drawingNode.getAttribute('r:id')));
      if (relation && workbook.files[relation.path]) {
        const drawingSource = parseXml(workbook.files[relation.path]);
        const rels = readRels(workbook.files, relation.path);
        for (const anchor of children(drawingSource.documentElement)) {
          const from = first(anchor, 'from');
          const rowNode = from && first(from, 'row');
          const absolute = anchor.localName === 'absoluteAnchor' ? positionFromAbsolute(anchor, sheet.doc.documentElement) : null;
          const oldRow = rowNode ? Number(rowNode.textContent) + 1 : absolute?.row || null;
          if (oldRow && !rowMap.has(oldRow)) continue;
          if (!all(anchor, 'pic').length) { warnings.add('原表中的非图片绘图对象未导出。'); continue; }
          let clone = drawingDoc.importNode(anchor, true);
          if (absolute) {
            clone = el(drawingDoc, 'xdr:oneCellAnchor', {}, NS.drawing);
            const newFrom = clone.appendChild(el(drawingDoc, 'xdr:from', {}, NS.drawing));
            for (const [tag, value] of [['col', absolute.col - 1], ['colOff', absolute.x], ['row', rowMap.get(oldRow) - 1], ['rowOff', absolute.y]]) {
              const node = newFrom.appendChild(el(drawingDoc, `xdr:${tag}`, {}, NS.drawing)); node.textContent = String(value);
            }
            for (const child of children(anchor)) if (child.localName !== 'pos') clone.appendChild(drawingDoc.importNode(child, true));
          } else if (oldRow) first(first(clone, 'from'), 'row').textContent = String(rowMap.get(oldRow) - 1);
          const to = first(clone, 'to');
          if (to) {
            const endRow = first(to, 'row'); const oldEnd = Number(endRow.textContent) + 1;
            const preceding = [...rowMap.keys()].filter(r => r < oldEnd).at(-1);
            endRow.textContent = String(rowMap.has(oldEnd) ? rowMap.get(oldEnd) - 1 : preceding ? rowMap.get(preceding) : 0);
          }
          for (const info of all(clone, 'cNvPr')) info.setAttribute('id', ++imageCount);
          const migrated = new Map();
          for (const node of [clone, ...all(clone, '*')]) {
            for (const attr of Array.from(node.attributes || [])) {
              if (attr.namespaceURI !== NS.rel && !attr.name.startsWith('r:')) continue;
              const id = attr.value;
              const rel = rels.find(r => r.id === id);
              if (!rel) throw new Error('原表图片的内部引用不完整，请先在 Excel / WPS 中重新保存。');
              if (migrated.has(id)) { node.setAttributeNS(NS.rel, attr.name, migrated.get(id)); continue; }
              const newId = `rImage${drawingRels.length + 1}`;
              if (rel.type.endsWith('/image') && !rel.external) {
                const target = copyMedia(workbook, rel.path);
                drawingRels.push({ id: newId, type: rel.type, target: `../media/${target.split('/').at(-1)}` });
              } else if (rel.type.endsWith('/hyperlink') && rel.external) {
                drawingRels.push({ ...rel, id: newId });
              } else if (/hlinkClick|hlinkHover/.test(node.localName)) {
                remove(node); warnings.add('图片上的内部跳转链接未迁移。'); continue;
              } else throw new Error('原表包含外链图片或不支持的图片附件，请先把图片嵌入原表。');
              migrated.set(id, newId);
              node.setAttributeNS(NS.rel, attr.name, newId);
            }
          }
          synchronizePictureGeometry(clone, root, anchor, sheet.doc.documentElement);
          drawingDoc.documentElement.appendChild(clone);
        }
      }
      remove(drawingNode);
    }
    for (const { ref, image } of pendingImages) {
      const target = copyMedia(workbook, image.path);
      const id = `rImage${drawingRels.length + 1}`;
      drawingRels.push({ id, type: `${NS.rel}/image`, target: `../media/${target.split('/').at(-1)}` });
      drawingDoc.documentElement.appendChild(makeCellAnchor(drawingDoc, image, workbook.files[image.path], ref, root, id, ++imageCount));
    }
    // Remove references to source-only features; the export is a self-contained selection document.
    for (const tag of ['legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'tableParts', 'extLst', 'rowBreaks', 'colBreaks', 'dataValidations', 'protectedRanges', 'sheetProtection', 'ignoredErrors', 'scenarios', 'customSheetViews', 'smartTags']) {
      for (const node of children(root, tag)) remove(node);
    }
    for (const node of all(root, 'pageSetup')) node.removeAttribute('r:id');
    for (const filter of children(root, 'autoFilter')) {
      const mapped = mapRange(filter.getAttribute('ref'), rowMap);
      if (mapped) filter.setAttribute('ref', mapped); else remove(filter);
    }
    let lastCol = 1;
    for (const cell of all(root, 'c')) lastCol = Math.max(lastCol, cellRef(cell.getAttribute('r'))?.column || 1);
    const dimension = first(root, 'dimension');
    dimension?.setAttribute('ref', `A1:${columnName(lastCol)}${rowMap.size}`);
    const views = first(root, 'sheetViews');
    for (const view of children(views)) {
      view.removeAttribute('tabSelected'); view.setAttribute('workbookViewId', 0); view.removeAttribute('topLeftCell');
      for (const node of children(view)) remove(node);
      view.appendChild(el(doc, 'selection', { activeCell: 'A1', sqref: 'A1' }));
    }
    if (children(drawingDoc.documentElement).length) {
      const drawingPath = `xl/drawings/drawing${number}.xml`;
      output[drawingPath] = xmlBytes(drawingDoc);
      output[relsPath(drawingPath)] = relationships(drawingRels);
      typeOverrides.push({ path: `/${drawingPath}`, type: `${CT}drawing+xml` });
      const node = el(doc, 'drawing'); node.setAttributeNS(NS.rel, 'r:id', 'rDrawing');
      root.appendChild(node);
      sheetRels.push({ id: 'rDrawing', type: `${NS.rel}/drawing`, target: `../drawings/drawing${number}.xml` });
    }
    const sheetPath = `xl/worksheets/sheet${number}.xml`;
    output[sheetPath] = xmlBytes(doc);
    if (sheetRels.length) output[relsPath(sheetPath)] = relationships(sheetRels);
    typeOverrides.push({ path: `/${sheetPath}`, type: `${CT}spreadsheetml.worksheet+xml` });
    const baseName = workbook.sheets.length > 1 ? `${workbook.name.replace(/\.xlsx$/i, '')}-${sheet.name}` : workbook.name;
    const tab = el(workbookDoc, 'sheet', { name: uniqueSheetName(baseName, usedNames), sheetId: number });
    tab.setAttributeNS(NS.rel, 'r:id', `rSheet${number}`); wbSheets.appendChild(tab);
    wbRels.push({ id: `rSheet${number}`, type: `${NS.rel}/worksheet`, target: `worksheets/sheet${number}.xml` });
  }
  if (nativeImages?.objectCount) {
    Object.assign(output, nativeImages.files());
    wbRels.push({ id: 'rCellImages', type: CELL_IMAGE_REL, target: 'cellimages.xml' });
    typeOverrides.push({ path: '/xl/cellimages.xml', type: CELL_IMAGE_TYPE });
  }
  output['xl/workbook.xml'] = xmlBytes(workbookDoc);
  output['xl/_rels/workbook.xml.rels'] = relationships(wbRels);
  output['_rels/.rels'] = relationships([{ id: 'rWorkbook', type: `${NS.rel}/officeDocument`, target: 'xl/workbook.xml' }]);
  const types = parseXml(`<Types xmlns="${NS.types}"/>`);
  for (const [extension, type] of [['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml'], ...mediaTypes]) types.documentElement.appendChild(el(types, 'Default', { Extension: extension, ContentType: type }, NS.types));
  for (const { path, type } of typeOverrides) types.documentElement.appendChild(el(types, 'Override', { PartName: path, ContentType: type }, NS.types));
  output['[Content_Types].xml'] = xmlBytes(types);
  const archive = Object.fromEntries(Object.entries(output).map(([path, bytes]) => [path, [bytes, { level: path.startsWith('xl/media/') ? 0 : 6 }]]));
  return { bytes: zipSync(archive), matchedCodes: result.matchedCodes, missingCodes: result.missingCodes, blockCount: result.blockCount, rowCount: result.rowCount, sheetCount: result.selections.length, imageCount, nativeImageCount: nativeImages?.count || 0, imageMode, warnings: [...warnings] };
}
