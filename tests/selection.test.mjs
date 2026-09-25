import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { zipSync, strToU8 } from 'fflate';
import { readWorkbook, parseSkuInput, matchSelection } from '../src/importer.js';
import { exportSelection } from '../src/exporter.js';
import { inspectXlsx, assertValidPackage, assertCopiedRows, assertCellImagesCopied, assertPictureGeometryMatchesAnchors, hash, first } from './xlsx-inspector.mjs';
import { makeFloatingImageWorkbook } from './synthetic-workbook.mjs';

const names = ['KOKO选款--8.5信息表(1).xlsx', 'Cole Haan-9.4选款.xlsx'];
const bytes = await Promise.all(names.map((name) => readFile(new URL(`../${name}`, import.meta.url))));
const initialHashes = bytes.map(hash);
const sources = bytes.map(inspectXlsx);
const books = await Promise.all(bytes.map((data, index) => readWorkbook(data, { name: names[index], id: `fixture-${index}` })));
const mapRows = (pairs) => new Map(pairs);

test('import recognizes both real files, merged SKU blocks, and numerical SKU', () => {
  assert.equal(books[0].skuCount, 7);
  assert.equal(books[0].rowCount, 16);
  assert.equal(books[1].skuCount, 18);
  assert.equal(books[1].rowCount, 37);
  const found = matchSelection(books, ['LR059', '18116', 'H140', 'DOES-NOT-EXIST']);
  assert.deepEqual(found.matchedCodes, ['LR059', '18116', 'H140']);
  assert.deepEqual(found.missingCodes, ['DOES-NOT-EXIST']);
  assert.equal(found.blockCount, 4);
  assert.equal(found.rowCount, 15);
});

test('SKU input accepts whitespace and Chinese separators, matches case-insensitively, and deduplicates', () => {
  const parsed = parseSkuInput(' lr059\nT026，H140;18116\tLR059、h140 ');
  assert.deepEqual(parsed.codes, ['LR059', 'T026', 'H140', '18116']);
  assert.equal(parsed.duplicates.length, 2);
  assert.deepEqual(parseSkuInput(' \n \t ').codes, []);
  const matched = matchSelection(books, parseSkuInput('lr059 h140').codes);
  assert.equal(matched.blockCount, 3);
  assert.equal(matched.rowCount, 13);
  assert.equal(matchSelection(books, ['LR05']).blockCount, 0, 'matching does not select partial SKUs');
});

test('LR059 copies the exact two color rows, complete A:T data, styles, dimensions, merges, and three images', async () => {
  const result = await exportSelection(books, ['LR059']);
  assert.equal(result.blockCount, 1);
  assert.equal(result.rowCount, 2);
  assert.equal(result.sheetCount, 1);
  assert.deepEqual(result.missingCodes, []);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  assert.equal(output.sheets.length, 1);
  const sheet = output.sheets[0];
  const rowMap = mapRows([[1, 1], [2, 2], [3, 3]]);
  assertCopiedRows(sources[0].sheets[0], sheet, rowMap);
  assertCellImagesCopied(sources[0], sources[0].sheets[0], sheet, rowMap);
  assert.equal(sheet.value('B2'), 'LR059');
  assert.equal(sheet.value('F2'), '米色');
  assert.equal(sheet.value('F3'), '深灰色');
  assert.equal(sheet.value('M2'), '32');
  assert.equal(sheet.value('M3'), '40');
  assert.equal(sheet.value('P2'), '460');
  assert.ok(sheet.merges.includes('B2:B3'));
  assert.ok(sheet.merges.includes('T2:T3'));
});

test('nonadjacent selections are compacted in original worksheet order without bringing in intervening products', async () => {
  const result = await exportSelection(books, ['H037', 'LR059', 'T041']);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  assert.equal(result.rowCount, 5);
  const rowMap = mapRows([[1, 1], [2, 2], [3, 3], [5, 4], [6, 5], [7, 6]]);
  assertCopiedRows(sources[0].sheets[0], output.sheets[0], rowMap);
  assertCellImagesCopied(sources[0], sources[0].sheets[0], output.sheets[0], rowMap);
  assert.equal(output.sheets[0].value('B2'), 'LR059');
  assert.equal(output.sheets[0].value('B4'), 'H037');
  assert.equal(output.sheets[0].value('B6'), 'T041');
});

test('converted WPS picture geometry follows cell anchors after compaction, preventing top-left image overlap', async () => {
  const output = inspectXlsx((await exportSelection(books, ['LR059', 'H037', 'T041'], { imageMode: 'drawing' })).bytes);
  const sheet = output.sheets[0];
  // These real fixtures use Normal = Microsoft YaHei 11 pt, whose digit width is 9 px.
  assertPictureGeometryMatchesAnchors(sheet, { maximumDigitWidth: 9 });
  const pictureAt = (row, col) => sheet.drawings.find((picture) => picture.row === row - 1 && picture.col === col - 1);
  const offset = (picture) => first(first(first(picture.anchor, 'spPr'), 'xfrm'), 'off');
  const firstColor = offset(pictureAt(2, 3));
  const secondColor = offset(pictureAt(3, 3));
  assert.ok(Number(firstColor.getAttribute('x')) > 0, 'LR059 picture begins in column C, not at the left edge');
  assert.ok(Number(firstColor.getAttribute('y')) > 0, 'LR059 picture begins beneath the header, not at the top edge');
  assert.ok(Number(secondColor.getAttribute('y')) > Number(firstColor.getAttribute('y')), 'second color picture is drawn below the first');
  assert.equal(sheet.value('B4'), 'H037', 'omitting T026 moves H037 from source row 5 to export row 4');
  const moved = pictureAt(4, 3);
  assert.ok(moved, 'H037 first picture moves with its row');
  const withinRow = Number(first(first(moved.anchor, 'from'), 'rowOff').textContent);
  assert.equal(Number(offset(moved).getAttribute('y')), (25.5 + 85 + 85) * 12700 + withinRow, 'compacted picture geometry excludes the omitted product row');
});

test('H140 from both sources remains separate, retaining every color and both original prices', async () => {
  const result = await exportSelection(books, ['H140']);
  assert.equal(result.sheetCount, 2);
  assert.equal(result.blockCount, 2);
  assert.equal(result.rowCount, 11);
  assert.deepEqual(result.matchedCodes, ['H140']);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  assert.equal(new Set(output.sheets.map((s) => s.name)).size, 2);
  const koko = output.sheets.find((s) => s.value('A1') === '序号');
  const cole = output.sheets.find((s) => s.value('A1') === 'PAMIR产品表');
  assert.ok(koko && cole, 'source-specific headers are kept separately');
  const kokoMap = mapRows([[1, 1], ...Array.from({ length: 6 }, (_, i) => [12 + i, 2 + i])]);
  const coleMap = mapRows([[1, 1], [2, 2], ...Array.from({ length: 5 }, (_, i) => [7 + i, 3 + i])]);
  assertCopiedRows(sources[0].sheets[0], koko, kokoMap);
  assertCopiedRows(sources[1].sheets[0], cole, coleMap);
  assertCellImagesCopied(sources[0], sources[0].sheets[0], koko, kokoMap);
  assertCellImagesCopied(sources[1], sources[1].sheets[0], cole, coleMap);
  assert.equal(koko.value('P2'), '353');
  assert.equal(cole.value('O3'), '477');
  assert.equal(cole.value('F6'), '深灰');
  assert.equal(cole.value('F7'), '深灰');
  assert.equal(cole.value('I6'), '49');
  assert.equal(cole.value('I7'), '15');
});

test('numeric SKU retains merged inventory and missing codes are reported with successful export', async () => {
  const result = await exportSelection(books, ['18116', 'NOT-FOUND']);
  assert.deepEqual(result.missingCodes, ['NOT-FOUND']);
  assert.deepEqual(result.matchedCodes, ['18116']);
  const output = inspectXlsx(result.bytes);
  const rowMap = mapRows([[1, 1], [10, 2], [11, 3]]);
  assertCopiedRows(sources[0].sheets[0], output.sheets[0], rowMap);
  assertCellImagesCopied(sources[0], sources[0].sheets[0], output.sheets[0], rowMap);
  assert.ok(output.sheets[0].merges.includes('G2:L2'));
  assert.equal(output.sheets[0].value('B2'), '18116');
  assert.equal(output.sheets[0].value('G2'), '122');
  assert.equal(output.sheets[0].value('G3'), '185');
});

test('all products survive a combined export, with collision-free styles and image relationships', async () => {
  const codes = ['LR059', 'T026', 'H037', 'T041', 'PZ032', '18116', 'H140', 'LR228', 'LR229', 'PZ930', 'HPZ887', 'PZ647', 'PZ754', 'PZ944', 'PZ945', 'HTF022', 'HH141', 'T089', 'HLR139', 'LR346', 'HPZ691', 'PZ347', 'PZ348', 'PZ349'];
  const result = await exportSelection(books, codes);
  assert.deepEqual(result.missingCodes, []);
  assert.equal(result.blockCount, 25);
  assert.equal(result.rowCount, 53);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  const koko = output.sheets.find((s) => s.value('A1') === '序号');
  const cole = output.sheets.find((s) => s.value('A1') === 'PAMIR产品表');
  const kokoMap = mapRows(Array.from({ length: 17 }, (_, i) => [i + 1, i + 1]));
  const coleOrder = Array.from({ length: 39 }, (_, i) => i + 1);
  const coleMap = mapRows(coleOrder.map((row, i) => [row, i + 1]));
  assertCopiedRows(sources[0].sheets[0], koko, kokoMap);
  assertCopiedRows(sources[1].sheets[0], cole, coleMap);
  assertCellImagesCopied(sources[0], sources[0].sheets[0], koko, kokoMap);
  assertCellImagesCopied(sources[1], sources[1].sheets[0], cole, coleMap);
});

test('invalid workbook and selections fail clearly instead of creating a misleading empty file', async () => {
  await assert.rejects(async () => readWorkbook(strToU8('not an Excel workbook'), { name: 'broken.xlsx', id: 'bad' }));
  const zip = zipSync({ 'unrelated.txt': strToU8('hello') });
  await assert.rejects(async () => readWorkbook(zip, { name: 'not-a-workbook.xlsx', id: 'bad-zip' }));
  await assert.rejects(async () => exportSelection(books, []));
  await assert.rejects(async () => exportSelection(books, ['NO-SUCH-SKU']));
});

test('ordinary Excel floating pictures follow selected rows, including bottom-boundary two-cell anchors', async () => {
  const picture = sources[0].files['xl/media/image1.jpeg'];
  const fixture = makeFloatingImageWorkbook(picture);
  const source = inspectXlsx(fixture);
  const workbook = await readWorkbook(fixture, { name: 'floating.xlsx', id: 'floating' });
  assert.equal(workbook.skuCount, 2);
  const result = await exportSelection([workbook], ['SECOND']);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  const sheet = output.sheets[0];
  assert.equal(result.rowCount, 2);
  assert.equal(sheet.value('B2'), 'SECOND');
  assert.equal(sheet.value('A2'), '黑色');
  assert.equal(sheet.value('A3'), '驼色');
  assert.ok(sheet.merges.includes('B2:B3'));
  assert.equal(sheet.drawings.length, 2, 'unselected FIRST picture is excluded');
  assertPictureGeometryMatchesAnchors(sheet, { checkHorizontal: false });
  for (const exported of sheet.drawings) {
    const name = first(exported.anchor, 'cNvPr').getAttribute('name');
    const original = source.sheets[0].drawings.find((item) => first(item.anchor, 'cNvPr').getAttribute('name') === name);
    assert.ok(original, 'exported picture corresponds to an original picture');
    const sourceTransform = first(first(original.anchor, 'spPr'), 'xfrm');
    const outputTransform = first(first(exported.anchor, 'spPr'), 'xfrm');
    assert.equal(first(outputTransform, 'off').getAttribute('x'), first(sourceTransform, 'off').getAttribute('x'), 'native image retains its horizontal geometry because columns did not move');
    assert.equal(first(outputTransform, 'ext').getAttribute('cx'), first(sourceTransform, 'ext').getAttribute('cx'), 'native image retains its original width');
  }
  assert.deepEqual(sheet.drawings.map((d) => [d.row, d.col]).sort(), [[1, 2], [2, 3]]);
  assert.ok(sheet.drawings.every((d) => d.hash === hash(picture)), 'floating picture bytes are preserved');
  const twoCell = sheet.drawings.find((d) => d.anchor.localName === 'twoCellAnchor');
  assert.ok(twoCell, 'existing two-cell positioning is retained');
  assert.equal(first(first(twoCell.anchor, 'to'), 'row').textContent, '3', 'bottom boundary moves together with both selected rows');
});

test('absolute-position pictures are filtered by their original row and shift with the selected product', async () => {
  const fixture = makeFloatingImageWorkbook(sources[0].files['xl/media/image1.jpeg'], { absolute: true });
  const workbook = await readWorkbook(fixture, { name: 'absolute.xlsx', id: 'absolute' });
  const output = inspectXlsx((await exportSelection([workbook], ['SECOND'])).bytes);
  assertValidPackage(output);
  const pictures = output.sheets[0].drawings;
  assert.equal(pictures.length, 2, 'FIRST absolute-position picture must be excluded');
  const positions = pictures.map((picture) => {
    const pos = first(picture.anchor, 'pos');
    if (pos) return [Number(pos.getAttribute('x')), Number(pos.getAttribute('y'))];
    const from = first(picture.anchor, 'from');
    assert.ok(from, 'absolute image has a valid absolute or converted cell anchor');
    const row = Number(first(from, 'row').textContent);
    const col = Number(first(from, 'col').textContent);
    return [col * 140 * 9525 + Number(first(from, 'colOff').textContent), (row ? 15 * 12700 + (row - 1) * 85 * 12700 : 0) + Number(first(from, 'rowOff').textContent)];
  }).sort(([a], [b]) => a - b);
  assert.deepEqual(positions, [[2667000, 285750], [4000500, 1365250]], 'remaining images move upward by the omitted two rows and keep their horizontal position and within-row offset');
});

test('image click hyperlinks retain their external relationship without dangling r:id', async () => {
  const fixture = makeFloatingImageWorkbook(sources[0].files['xl/media/image1.jpeg'], { hyperlinks: true });
  const workbook = await readWorkbook(fixture, { name: 'picture-links.xlsx', id: 'picture-links' });
  const output = inspectXlsx((await exportSelection([workbook], ['SECOND'])).bytes);
  assertValidPackage(output);
  assert.equal(output.sheets[0].drawings.length, 2);
  for (const picture of output.sheets[0].drawings) {
    const click = first(picture.anchor, 'hlinkClick');
    assert.ok(click, 'picture click action remains present');
    const rel = output.rels(picture.part).get(click.getAttribute('r:id'));
    assert.ok(rel, 'picture click action resolves to an exported relationship');
    assert.equal(rel.external, true);
    assert.match(rel.type, /\/hyperlink$/);
    assert.equal(rel.target, 'https://example.com/products/second?color=black&size=M');
  }
});

test('SVG image extension and bitmap fallback both retain embedded bytes and valid relationships', async () => {
  const fixture = makeFloatingImageWorkbook(sources[0].files['xl/media/image1.jpeg'], { svg: true });
  const source = inspectXlsx(fixture);
  const workbook = await readWorkbook(fixture, { name: 'vector.xlsx', id: 'vector' });
  const output = inspectXlsx((await exportSelection([workbook], ['SECOND'])).bytes);
  assertValidPackage(output);
  assert.equal(output.sheets[0].drawings.length, 2);
  for (const picture of output.sheets[0].drawings) {
    const svg = first(picture.anchor, 'svgBlip');
    assert.ok(svg, 'SVG extension remains present');
    const rel = output.rels(picture.part).get(svg.getAttribute('r:embed'));
    assert.ok(rel && output.files[rel.path], 'SVG embed resolves to exported media');
    assert.equal(hash(output.files[rel.path]), hash(source.files['xl/media/vector.svg']), 'SVG bytes are preserved');
    assert.equal(picture.hash, hash(sources[0].files['xl/media/image1.jpeg']), 'bitmap fallback bytes are preserved');
  }
});

test('import and repeated exports leave original workbook bytes and input files untouched', async () => {
  await exportSelection(books, ['LR059']);
  await exportSelection(books, ['H140']);
  assert.deepEqual(bytes.map(hash), initialHashes, 'input byte buffers are unchanged');
  const disk = await Promise.all(names.map((name) => readFile(new URL(`../${name}`, import.meta.url))));
  assert.deepEqual(disk.map(hash), initialHashes, 'original Excel files are unchanged');
});
