import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

// Independent OOXML inspection. No application parsing/export code is reused.
export const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
export const all = (node, name) => Array.from(node?.getElementsByTagNameNS('*', name) || []);
export const first = (node, name) => all(node, name)[0];
export const children = (node) => Array.from(node?.childNodes || []).filter((n) => n.nodeType === 1);
export const attrs = (node) => Object.fromEntries(Array.from(node?.attributes || []).filter((a) => !a.name.startsWith('xmlns')).map((a) => [a.name, a.value]).sort(([a], [b]) => a.localeCompare(b)));
export function canonical(node) {
  if (!node) return null;
  return { name: node.localName, attrs: attrs(node), children: Array.from(node.childNodes).filter((n) => n.nodeType === 1 || (n.nodeType === 3 && n.nodeValue.trim())).map((n) => n.nodeType === 1 ? canonical(n) : n.nodeValue) };
}

// Picture identities and package relationship numbers can change when sources
// are combined. Every visual property and extension must otherwise survive.
export function canonicalPicture(node) {
  const copy = node.cloneNode(true);
  for (const item of [copy, ...all(copy, '*')]) {
    if (item.localName === 'cNvPr') {
      item.removeAttribute('id');
      item.removeAttribute('name');
    }
    for (const attribute of Array.from(item.attributes || [])) {
      if (attribute.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') item.setAttributeNS(attribute.namespaceURI, attribute.name, 'RELATIONSHIP');
    }
  }
  return canonical(copy);
}

export const cellImageId = (formula) => formula.match(/DISPIMG\(\s*"([^"\n]+)"/i)?.[1];

export function parseXml(bytes, label = 'XML') {
  assert.ok(bytes, `${label} exists`);
  const failures = [];
  const xml = new DOMParser({ errorHandler: { warning: (m) => failures.push(m), error: (m) => failures.push(m), fatalError: (m) => failures.push(m) } }).parseFromString(typeof bytes === 'string' ? bytes : strFromU8(bytes), 'application/xml');
  assert.deepEqual(failures, [], `${label} is well formed`);
  assert.ok(xml.documentElement, `${label} has a root element`);
  assert.equal(all(xml, 'parsererror').length, 0, `${label} contains no parse errors`);
  return xml;
}

const relPath = (part) => path.posix.join(path.posix.dirname(part), '_rels', `${path.posix.basename(part)}.rels`);
const resolve = (part, target) => target.startsWith('/') ? target.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(part), target));
export function inspectXlsx(bytes) {
  const files = unzipSync(bytes);
  const xml = (part) => parseXml(files[part], part);
  const rels = (part) => files[relPath(part)] ? new Map(all(xml(relPath(part)), 'Relationship').map((r) => [r.getAttribute('Id'), { path: resolve(part, r.getAttribute('Target')), target: r.getAttribute('Target'), type: r.getAttribute('Type'), external: r.getAttribute('TargetMode') === 'External' }])) : new Map();
  const strings = files['xl/sharedStrings.xml'] ? all(xml('xl/sharedStrings.xml'), 'si').map((si) => all(si, 't').map((t) => t.textContent).join('')) : [];
  const styles = files['xl/styles.xml'] ? xml('xl/styles.xml') : null;
  const themePart = Object.keys(files).find((part) => /^xl\/theme\/theme\d+\.xml$/.test(part));
  const colorScheme = themePart ? first(xml(themePart), 'clrScheme') : null;
  const themeNames = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const themeColors = themeNames.map((name) => {
    const color = children(children(colorScheme).find((n) => n.localName === name))[0];
    return color ? color.getAttribute('lastClr') || color.getAttribute('val') : null;
  });
  const canonicalStyle = (node) => {
    const result = canonical(node);
    const normalize = (item) => {
      if (!item || typeof item === 'string') return;
      if (Object.hasOwn(item.attrs, 'theme') && themeColors[Number(item.attrs.theme)]) {
        item.attrs.rgb = `FF${themeColors[Number(item.attrs.theme)]}`.toUpperCase();
        delete item.attrs.theme;
      }
      item.children.forEach(normalize);
    };
    normalize(result);
    return result;
  };
  const cellValue = (c) => {
    if (!c) return '';
    const value = first(c, 'v')?.textContent || '';
    if (c.getAttribute('t') === 's') return strings[Number(value)];
    if (c.getAttribute('t') === 'inlineStr') return all(c, 't').map((t) => t.textContent).join('');
    return value;
  };
  const style = (c) => {
    if (!styles) return null;
    const xf = children(first(styles, 'cellXfs'))[Number(c?.getAttribute('s') || 0)];
    assert.ok(xf, `style ${c?.getAttribute('s')} exists`);
    const normalized = canonicalStyle(xf);
    for (const [attribute, collection] of [['fontId', 'fonts'], ['fillId', 'fills'], ['borderId', 'borders']]) {
      normalized.attrs[attribute] = canonicalStyle(children(first(styles, collection))[Number(xf.getAttribute(attribute) || 0)]);
    }
    const numFmt = all(styles, 'numFmt').find((n) => n.getAttribute('numFmtId') === xf.getAttribute('numFmtId'));
    normalized.attrs.numFmtId = numFmt?.getAttribute('formatCode') ?? xf.getAttribute('numFmtId');
    const base = children(first(styles, 'cellStyleXfs'))[Number(xf.getAttribute('xfId') || 0)];
    // Parent IDs may be rebased when style collections are combined; compare their content.
    if (base) {
      normalized.attrs.xfId = canonicalStyle(base);
      for (const [attribute, collection] of [['fontId', 'fonts'], ['fillId', 'fills'], ['borderId', 'borders']]) normalized.attrs.xfId.attrs[attribute] = canonicalStyle(children(first(styles, collection))[Number(base.getAttribute(attribute) || 0)]);
      const baseNumFmt = all(styles, 'numFmt').find((n) => n.getAttribute('numFmtId') === base.getAttribute('numFmtId'));
      normalized.attrs.xfId.attrs.numFmtId = baseNumFmt?.getAttribute('formatCode') ?? base.getAttribute('numFmtId');
    }
    return normalized;
  };
  const workbookRels = rels('xl/workbook.xml');
  const sheets = all(xml('xl/workbook.xml'), 'sheet').map((item) => {
    const part = workbookRels.get(item.getAttribute('r:id')).path;
    const doc = xml(part);
    const rows = new Map(all(doc, 'row').map((r) => [Number(r.getAttribute('r')), r]));
    const cells = new Map(all(doc, 'c').map((c) => [c.getAttribute('r'), c]));
    const sheetRels = rels(part);
    const drawings = [];
    for (const drawing of all(doc, 'drawing')) {
      const drawingPart = sheetRels.get(drawing.getAttribute('r:id'))?.path;
      assert.ok(drawingPart && files[drawingPart], `drawing for ${part} resolves`);
      const drawingRels = rels(drawingPart);
      for (const anchor of children(xml(drawingPart).documentElement)) {
        const from = first(anchor, 'from');
        for (const blip of all(anchor, 'blip')) {
          const media = drawingRels.get(blip.getAttribute('r:embed'))?.path;
          assert.ok(media && files[media], `drawing media resolves from ${drawingPart}`);
          drawings.push({ row: Number(first(from, 'row')?.textContent), col: Number(first(from, 'col')?.textContent), hash: hash(files[media]), anchor, media, part: drawingPart });
        }
      }
    }
    return { name: item.getAttribute('name'), part, doc, rows, cells, drawings, merges: all(doc, 'mergeCell').map((m) => m.getAttribute('ref')), value: (ref) => cellValue(cells.get(ref)), style: (ref) => style(cells.get(ref)) };
  });
  const cellImages = new Map();
  if (files['xl/cellimages.xml']) {
    const imageRels = rels('xl/cellimages.xml');
    for (const pic of all(xml('xl/cellimages.xml'), 'pic')) {
      const name = first(pic, 'cNvPr').getAttribute('name');
      const media = imageRels.get(first(pic, 'blip').getAttribute('r:embed'))?.path;
      assert.ok(files[media], `cell image ${name} resolves`);
      assert.ok(!cellImages.has(name), `cell image ${name} has a unique identifier`);
      cellImages.set(name, { hash: hash(files[media]), media, pic, item: pic.parentNode, part: 'xl/cellimages.xml' });
    }
  }
  for (const sheet of sheets) {
    sheet.cellPictures = [];
    for (const [ref, cell] of sheet.cells) {
      const formula = first(cell, 'f')?.textContent || first(cell, 'v')?.textContent || '';
      const id = cellImageId(formula);
      if (!id) continue;
      const picture = cellImages.get(id);
      assert.ok(picture, `cell image at ${sheet.name}!${ref} resolves to ${id}`);
      const letters = ref.match(/^[A-Z]+/)[0];
      const col = [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      sheet.cellPictures.push({ ...picture, id, ref, formula, row: Number(ref.match(/\d+$/)[0]) - 1, col });
    }
  }
  return { files, xml, rels, sheets, cellImages };
}

export function assertValidPackage(book) {
  assert.ok(book.files['[Content_Types].xml']);
  for (const [part, bytes] of Object.entries(book.files)) {
    if (!part.endsWith('.xml') && !part.endsWith('.rels')) continue;
    const doc = parseXml(bytes, part);
    for (const node of all(doc, '*')) {
      if (node.prefix) assert.ok(node.namespaceURI, `${part}: prefix ${node.prefix} is declared`);
      if (part.endsWith('.xml')) {
        for (const attribute of Array.from(node.attributes || [])) {
          if (attribute.namespaceURI !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') continue;
          assert.ok(book.rels(part).has(attribute.value), `${part}: ${node.nodeName} ${attribute.name} resolves to an existing relationship`);
        }
      }
    }
    if (!part.endsWith('.rels')) continue;
    const owner = part === '_rels/.rels' ? '' : part.replace(/\/\_rels\/([^/]+)\.rels$/, '/$1');
    for (const rel of all(doc, 'Relationship')) {
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const target = resolve(owner, rel.getAttribute('Target'));
      assert.ok(book.files[target], `${part} references existing ${target}`);
    }
  }
  for (const type of all(book.xml('[Content_Types].xml'), 'Override')) assert.ok(book.files[type.getAttribute('PartName').replace(/^\//, '')], `content type references existing ${type.getAttribute('PartName')}`);
}

export function assertCopiedRows(source, target, rowMap) {
  const expectedRows = [...rowMap.values()].sort((a, b) => a - b);
  assert.deepEqual([...target.rows.keys()], expectedRows, 'only the chosen full rows and original headers are present, without gaps');
  for (const [oldRow, newRow] of rowMap) {
    for (const name of ['ht', 'customHeight', 'hidden', 'outlineLevel']) assert.equal(target.rows.get(newRow).getAttribute(name), source.rows.get(oldRow).getAttribute(name), `row ${oldRow} -> ${newRow} preserves ${name}`);
    const oldCells = children(source.rows.get(oldRow)).filter((c) => c.localName === 'c');
    for (const cell of oldCells) {
      const oldRef = cell.getAttribute('r');
      const newRef = oldRef.replace(/\d+$/, String(newRow));
      const formula = first(cell, 'f')?.textContent || '';
      if (!formula.includes('DISPIMG')) assert.equal(target.value(newRef), source.value(oldRef), `complete row data ${oldRef} -> ${newRef}`);
      assert.deepEqual(target.style(newRef), source.style(oldRef), `cell style ${oldRef} -> ${newRef}`);
    }
  }
  const expectedMerges = source.merges.filter((ref) => {
    const [from, to = from] = ref.split(':').map((c) => Number(c.match(/\d+$/)[0]));
    return Array.from({ length: to - from + 1 }, (_, i) => from + i).every((r) => rowMap.has(r));
  }).map((ref) => ref.replace(/\d+/g, (r) => String(rowMap.get(Number(r))))).sort();
  assert.deepEqual([...target.merges].sort(), expectedMerges, 'all complete selected merges are preserved and moved with their rows');
  const sourceCols = all(source.doc, 'col');
  const targetCols = all(target.doc, 'col');
  for (let column = 1; column <= Math.max(...sourceCols.map((c) => Number(c.getAttribute('max')))); column++) {
    const before = sourceCols.find((c) => +c.getAttribute('min') <= column && +c.getAttribute('max') >= column);
    const after = targetCols.find((c) => +c.getAttribute('min') <= column && +c.getAttribute('max') >= column);
    for (const prop of ['width', 'customWidth', 'hidden', 'bestFit']) assert.equal(after?.getAttribute(prop), before?.getAttribute(prop), `column ${column} preserves ${prop}`);
  }
}

export function assertCellImagesCopied(sourceBook, sourceSheet, targetSheet, rowMap) {
  const expected = [];
  for (const [oldRow, newRow] of rowMap) {
    for (const cell of children(sourceSheet.rows.get(oldRow)).filter((c) => c.localName === 'c')) {
      const formula = first(cell, 'f')?.textContent || '';
      const id = cellImageId(formula);
      if (!id) continue;
      const ref = cell.getAttribute('r');
      const letters = ref.match(/^[A-Z]+/)[0];
      const col = [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
      const merge = sourceSheet.merges.find((range) => range.split(':')[0] === ref);
      const lastRow = merge ? Number(merge.split(':')[1].match(/\d+$/)[0]) : oldRow;
      expected.push({ ...sourceBook.cellImages.get(id), id, formula, firstRow: newRow - 1, lastRow: rowMap.get(lastRow) - 1, col, ref, newRef: `${letters}${newRow}` });
    }
  }
  const unmatched = [...targetSheet.drawings, ...targetSheet.cellPictures];
  assert.equal(unmatched.length, expected.length, 'all selected product and size-table images, without unrelated images');
  for (const item of expected) {
    const index = unmatched.findIndex((d) => d.hash === item.hash && d.col === item.col && d.row >= item.firstRow && d.row <= item.lastRow);
    assert.ok(index >= 0, `image at ${item.ref} retains exact media bytes and belongs to the correct cell/merged range`);
    const exported = unmatched[index];
    if (exported.ref) {
      assert.equal(exported.ref, item.newRef, `native image remains bound to the exact copied cell ${item.newRef}`);
      assert.equal(exported.formula, item.formula.replace(item.id, exported.id), 'DISPIMG formula and display mode are preserved');
      assert.deepEqual(canonicalPicture(exported.item), canonicalPicture(item.item), 'native image metadata, dimensions, crop and transforms are preserved');
    }
    unmatched.splice(index, 1);
  }
}

export function assertPictureGeometryMatchesAnchors(sheet, { maximumDigitWidth = 7, checkHorizontal = true } = {}) {
  const defaults = first(sheet.doc, 'sheetFormatPr');
  const columnDefinitions = all(sheet.doc, 'col');
  const point = (marker) => {
    const columnCount = Number(first(marker, 'col').textContent);
    const rowCount = Number(first(marker, 'row').textContent);
    let x = Number(first(marker, 'colOff').textContent);
    let y = Number(first(marker, 'rowOff').textContent);
    for (let column = 1; column <= columnCount; column++) {
      const definition = columnDefinitions.find((c) => Number(c.getAttribute('min')) <= column && Number(c.getAttribute('max')) >= column);
      if (definition?.getAttribute('hidden') === '1') continue;
      const width = Number(definition?.getAttribute('width') || defaults?.getAttribute('defaultColWidth') || 8.43);
      // OOXML column width already includes padding; use the Normal font's digit width.
      x += Math.floor((width * 256 + Math.trunc(128 / maximumDigitWidth)) * maximumDigitWidth / 256) * 9525;
    }
    for (let row = 1; row <= rowCount; row++) {
      const definition = sheet.rows.get(row);
      if (definition?.getAttribute('hidden') === '1') continue;
      y += Number(definition?.getAttribute('ht') || defaults?.getAttribute('defaultRowHeight') || 15) * 12700;
    }
    return { x: Math.round(x), y: Math.round(y) };
  };
  for (const picture of sheet.drawings) {
    const from = first(picture.anchor, 'from');
    if (!from) continue;
    const expected = point(from);
    const transform = first(first(picture.anchor, 'spPr'), 'xfrm');
    assert.ok(transform, `picture at row ${picture.row + 1}, column ${picture.col + 1} has explicit geometry`);
    const offset = first(transform, 'off');
    if (checkHorizontal) assert.equal(Number(offset.getAttribute('x')), expected.x, 'picture geometry and cell anchor have the same absolute horizontal position');
    assert.equal(Number(offset.getAttribute('y')), expected.y, 'picture geometry and cell anchor have the same absolute vertical position');
    const to = first(picture.anchor, 'to');
    const endpoint = to && point(to);
    const anchorExtent = first(picture.anchor, 'ext');
    const extent = first(transform, 'ext');
    if (checkHorizontal) assert.equal(Number(extent.getAttribute('cx')), endpoint ? endpoint.x - expected.x : Number(anchorExtent.getAttribute('cx')), 'picture geometry width matches anchor width');
    assert.equal(Number(extent.getAttribute('cy')), endpoint ? endpoint.y - expected.y : Number(anchorExtent.getAttribute('cy')), 'picture geometry height matches anchor height');
  }
}
