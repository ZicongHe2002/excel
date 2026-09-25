import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { readWorkbook } from '../src/importer.js';
import { exportSelection } from '../src/exporter.js';
import { inspectXlsx, assertValidPackage, assertCellImagesCopied, canonicalPicture, cellImageId, first, all, parseXml, hash } from './xlsx-inspector.mjs';

const names = ['KOKO选款--8.5信息表(1).xlsx', 'Cole Haan-9.4选款.xlsx'];
const bytes = await Promise.all(names.map((name) => readFile(new URL(`../${name}`, import.meta.url))));
const sources = bytes.map(inspectXlsx);
const workbooks = bytes.map((data, i) => readWorkbook(data, { name: names[i], id: `native-${i}` }));
const imageAt = (book, ref, sheet = 0) => book.sheets[sheet].cellPictures.find((image) => image.ref === ref);
const mediaFiles = (book) => Object.keys(book.files).filter((name) => /^xl\/media\//.test(name));

function modifiedKoko(mutator) {
  const files = unzipSync(bytes[0]);
  const documents = new Map();
  const xml = (part) => {
    if (!documents.has(part)) documents.set(part, parseXml(files[part], part));
    return documents.get(part);
  };
  mutator({ files, xml, source: sources[0] });
  for (const [part, doc] of documents) files[part] = strToU8(doc.toString());
  return zipSync(files);
}

function fixturePicture(xml, id) {
  return all(xml('xl/cellimages.xml'), 'pic').find((pic) => first(pic, 'cNvPr').getAttribute('name') === id);
}

function fixtureCell(xml, ref) {
  return all(xml('xl/worksheets/sheet1.xml'), 'c').find((cell) => cell.getAttribute('r') === ref);
}

test('all 60 real cell images retain native objects, raw bytes, formula modes and original cell bindings', async () => {
  const codes = workbooks.flatMap((book) => book.sheets.flatMap((sheet) => sheet.blocks.map((block) => block.normalizedSku)));
  const result = await exportSelection(workbooks, codes);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  assert.equal(result.imageMode, 'original');
  assert.equal(result.imageCount, 60);
  assert.equal(result.nativeImageCount, 60);
  assert.equal(output.cellImages.size, 60);
  assert.equal(output.sheets[0].cellPictures.length, 23);
  assert.equal(output.sheets[1].cellPictures.length, 37);
  assert.equal(output.sheets.flatMap((sheet) => sheet.drawings).length, 0, 'native cell images are not rebuilt as floating pictures');
  const expectedMedia = new Set();
  for (let i = 0; i < sources.length; i++) {
    const sourceSheet = sources[i].sheets[0];
    const targetSheet = output.sheets[i];
    const rowMap = new Map([...targetSheet.rows.keys()].map((row) => [row, row]));
    assertCellImagesCopied(sources[i], sourceSheet, targetSheet, rowMap);
    for (const sourceImage of sourceSheet.cellPictures) {
      const exported = imageAt(output, sourceImage.ref, i);
      assert.ok(exported, `native image ${sourceImage.ref} stays in its source sheet`);
      assert.deepEqual(output.files[exported.media], sources[i].files[sourceImage.media], `every byte of ${names[i]}!${sourceImage.ref} is unchanged`);
      assert.deepEqual(canonicalPicture(exported.item), canonicalPicture(sourceImage.item), 'entire cellImage object metadata survives package remapping');
      const sourceCell = sourceSheet.cells.get(sourceImage.ref);
      const targetCell = targetSheet.cells.get(sourceImage.ref);
      assert.equal(first(targetCell, 'v').textContent, first(sourceCell, 'v').textContent.replace(sourceImage.id, exported.id), 'cached DISPIMG expression stays synchronized with its formula');
      expectedMedia.add(sourceImage.hash);
    }
  }
  assert.deepEqual(new Set(mediaFiles(output).map((part) => hash(output.files[part]))), expectedMedia, 'output has precisely the selected original media');
});

test('LR059 keeps both color pictures and its size-chart as three unchanged native images without unselected media', async () => {
  const output = inspectXlsx((await exportSelection(workbooks, ['LR059'])).bytes);
  assertValidPackage(output);
  const sheet = output.sheets[0];
  assert.deepEqual(sheet.cellPictures.map((picture) => picture.ref).sort(), ['C2', 'C3', 'T2']);
  assert.equal(sheet.drawings.length, 0);
  assert.equal(output.cellImages.size, 3);
  assert.equal(mediaFiles(output).length, 3);
  for (const picture of sheet.cellPictures) {
    const original = imageAt(sources[0], picture.ref);
    assert.equal(picture.id, original.id, 'noncolliding image identifier is unchanged');
    assert.equal(picture.formula, original.formula, 'original DISPIMG mode 1 and formula are unchanged');
    assert.deepEqual(output.files[picture.media], sources[0].files[original.media]);
    assert.deepEqual(canonicalPicture(picture.item), canonicalPicture(original.item));
  }
  assert.ok(sheet.merges.includes('T2:T3'), 'size-chart remains in its original two-row merged cell');
  assert.ok(!Object.keys(output.files).some((part) => part.startsWith('xl/drawings/')), 'native export introduces no drawing parts');
});

test('same native image IDs in separate source files cannot substitute one source image for another', async () => {
  const firstImage = imageAt(sources[0], 'C2');
  const replacement = imageAt(sources[0], 'C3');
  assert.notEqual(firstImage.hash, replacement.hash);
  const modified = modifiedKoko(({ files }) => { files[firstImage.media] = files[replacement.media]; });
  const second = readWorkbook(modified, { name: 'same-ids-different-pictures.xlsx', id: 'native-collision' });
  const output = inspectXlsx((await exportSelection([workbooks[0], second], ['LR059'])).bytes);
  assertValidPackage(output);
  assert.equal(output.cellImages.size, 6);
  const originalExport = imageAt(output, 'C2', 0);
  const collisionExport = imageAt(output, 'C2', 1);
  assert.equal(originalExport.id, firstImage.id, 'first source keeps its existing identifier');
  assert.notEqual(collisionExport.id, firstImage.id, 'colliding identifier is reassigned in the second source');
  assert.equal(originalExport.hash, firstImage.hash);
  assert.equal(collisionExport.hash, replacement.hash, 'second source retains its own image bytes');
  assert.equal(cellImageId(first(output.sheets[1].cells.get('C2'), 'v').textContent), collisionExport.id, 'cached native image expression uses the reassigned identifier');
  assert.deepEqual(canonicalPicture(collisionExport.item), canonicalPicture(firstImage.item), 'collision resolution changes references, preserving original image geometry');
});

test('a native picture referenced by two cells is copied once and remains attached to both cells', async () => {
  const shared = imageAt(sources[0], 'C2');
  const discarded = imageAt(sources[0], 'C3');
  const modified = modifiedKoko(({ xml }) => {
    const cell = fixtureCell(xml, 'C3');
    for (const tag of ['f', 'v']) first(cell, tag).textContent = first(cell, tag).textContent.replace(discarded.id, shared.id);
  });
  const workbook = readWorkbook(modified, { name: 'shared-image.xlsx', id: 'native-shared' });
  const result = await exportSelection([workbook], ['LR059']);
  const output = inspectXlsx(result.bytes);
  assertValidPackage(output);
  assert.equal(result.imageCount, 3, 'three selected image-bearing cells');
  assert.equal(output.sheets[0].cellPictures.length, 3);
  assert.equal(output.cellImages.size, 2, 'one shared product image plus one size chart');
  assert.equal(mediaFiles(output).length, 2);
  assert.equal(imageAt(output, 'C2').id, imageAt(output, 'C3').id);
  assert.equal(imageAt(output, 'C2').hash, shared.hash);
  assert.ok(!output.cellImages.has(discarded.id), 'unreferenced native object is excluded');
});

test('native image crops, rotations, offsets, extents, transparency and mode 2 are copied without reconstruction', async () => {
  const original = imageAt(sources[0], 'C2');
  const modified = modifiedKoko(({ xml }) => {
    const picture = fixturePicture(xml, original.id);
    const transform = first(picture, 'xfrm');
    transform.setAttribute('rot', '5400000');
    transform.setAttribute('flipH', '1');
    first(transform, 'off').setAttribute('x', '12345');
    first(transform, 'off').setAttribute('y', '67890');
    first(transform, 'ext').setAttribute('cx', '2345678');
    first(transform, 'ext').setAttribute('cy', '3456789');
    const fill = first(picture, 'blipFill');
    const crop = picture.ownerDocument.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:srcRect');
    for (const [name, value] of Object.entries({ l: '10000', r: '2500', t: '4000', b: '6000' })) crop.setAttribute(name, value);
    fill.insertBefore(crop, first(fill, 'stretch'));
    const transparency = picture.ownerDocument.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:alphaModFix');
    transparency.setAttribute('amt', '80000');
    first(picture, 'blip').appendChild(transparency);
    const cell = fixtureCell(xml, 'C2');
    for (const tag of ['f', 'v']) first(cell, tag).textContent = first(cell, tag).textContent.replace(',1)', ',2)');
  });
  const source = inspectXlsx(modified);
  const workbook = readWorkbook(modified, { name: 'image-effects.xlsx', id: 'native-effects' });
  const output = inspectXlsx((await exportSelection([workbook], ['LR059'])).bytes);
  assertValidPackage(output);
  const exported = imageAt(output, 'C2');
  assert.equal(exported.formula, imageAt(source, 'C2').formula);
  assert.match(exported.formula, /,2\)$/);
  assert.deepEqual(canonicalPicture(exported.item), canonicalPicture(imageAt(source, 'C2').item));
  assert.deepEqual(output.files[exported.media], source.files[imageAt(source, 'C2').media]);
  assert.equal(output.sheets[0].drawings.length, 0);
});

test('a selected native image with a missing relationship or missing media fails clearly', async () => {
  const selected = imageAt(sources[0], 'C2');
  const mutations = [
    ({ xml }) => first(fixturePicture(xml, selected.id), 'blip').setAttribute('r:embed', 'rMissingImage'),
    ({ files }) => { delete files[selected.media]; },
  ];
  for (const [index, mutate] of mutations.entries()) {
    const workbook = readWorkbook(modifiedKoko(mutate), { name: 'broken-native-image.xlsx', id: `native-broken-${index}` });
    await assert.rejects(() => exportSelection([workbook], ['LR059']), /图片|原图/, 'missing selected image data must not silently create a partial export');
  }
});

test('unselected native image data is neither copied nor required by the selected product', async () => {
  const unrelated = imageAt(sources[0], 'C4');
  const modified = modifiedKoko(({ files }) => { delete files[unrelated.media]; });
  const workbook = readWorkbook(modified, { name: 'unselected-image-missing.xlsx', id: 'native-unselected' });
  const output = inspectXlsx((await exportSelection([workbook], ['LR059'])).bytes);
  assertValidPackage(output);
  assert.equal(output.cellImages.size, 3);
  assert.equal(mediaFiles(output).length, 3);
  assert.ok(!output.cellImages.has(unrelated.id));
});
