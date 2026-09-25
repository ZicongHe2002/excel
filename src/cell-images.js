import { NS, parseXml, all, children, first, el, readRels, relationships, xmlBytes } from './xml.js';

const CELL_IMAGE_NS = 'http://www.wps.cn/officeDocument/2017/etCustomData';
export const CELL_IMAGE_REL = 'http://www.wps.cn/officeDocument/2020/cellImage';
export const CELL_IMAGE_TYPE = 'application/vnd.wps-officedocument.cellimage+xml';

// The image object and its raw media are kept intact. Only package references
// and colliding object names change when combining independent workbooks.
export function createCellImageCopier(workbooks, copyMedia) {
  const catalogs = new Map();
  const copied = new Map();
  const usedNames = new Set();
  const rels = [];
  const doc = parseXml(`<etc:cellImages xmlns:etc="${CELL_IMAGE_NS}" xmlns:xdr="${NS.drawing}" xmlns:a="${NS.art}" xmlns:r="${NS.rel}"/>`);
  let nextName = 0;
  let nextPictureId = 0;
  let references = 0;

  for (const workbook of workbooks) {
    const catalog = new Map();
    const workbookRelations = readRels(workbook.files, workbook.workbookPath || 'xl/workbook.xml');
    const paths = new Set(workbookRelations.filter(r => r.type.endsWith('/cellImage') && !r.external).map(r => r.path));
    for (const path of Object.keys(workbook.files)) if (/(^|\/)cellimages\.xml$/i.test(path)) paths.add(path);
    for (const path of paths) {
      if (!workbook.files[path]) throw new Error(`${workbook.name} 缺少单元格图片数据，请重新保存原表。`);
      const source = parseXml(workbook.files[path]);
      const sourceRels = readRels(workbook.files, path);
      for (const item of children(source.documentElement, 'cellImage')) {
        const picture = all(item, 'pic')[0];
        const name = picture && all(picture, 'cNvPr')[0]?.getAttribute('name');
        if (!name) continue;
        if (catalog.has(name)) throw new Error(`${workbook.name} 含重复的图片标识，请在 WPS 中重新保存原表。`);
        catalog.set(name, { item, sourceRels, root: source.documentElement });
      }
    }
    catalogs.set(workbook, catalog);
    copied.set(workbook, new Map());
  }

  function copy(workbook, sourceName) {
    references++;
    const existing = copied.get(workbook).get(sourceName);
    if (existing) return existing;
    const source = catalogs.get(workbook).get(sourceName);
    if (!source) throw new Error(`${workbook.name} 缺少图片 ${sourceName}，请重新保存包含图片的原表。`);
    let name = sourceName;
    while (usedNames.has(name)) name = `ID_${(++nextName).toString(16).toUpperCase().padStart(32, '0')}`;
    usedNames.add(name);
    const item = doc.importNode(source.item, true);
    // Keep prefix declarations used by extension attributes or picture effects.
    for (const attr of Array.from(source.root.attributes || [])) {
      if (attr.name.startsWith('xmlns:') && !item.hasAttribute(attr.name)) item.setAttributeNS('http://www.w3.org/2000/xmlns/', attr.name, attr.value);
    }
    for (const info of all(item, 'cNvPr')) info.setAttribute('id', ++nextPictureId);
    all(item, 'cNvPr')[0].setAttribute('name', name);
    const relationMap = new Map();
    let embeddedCount = 0;
    for (const node of [item, ...all(item, '*')]) {
      for (const attr of Array.from(node.attributes || [])) {
        if (attr.namespaceURI !== NS.rel) continue;
        const sourceId = attr.value;
        if (relationMap.has(sourceId)) { node.setAttributeNS(NS.rel, attr.name, relationMap.get(sourceId)); continue; }
        const relation = source.sourceRels.find(r => r.id === sourceId);
        if (!relation) throw new Error(`${workbook.name} 的图片引用不完整，请在 WPS 中重新保存原表。`);
        const id = `rCellImage${rels.length + 1}`;
        if (relation.type.endsWith('/image') && !relation.external) {
          const target = copyMedia(workbook, relation.path);
          rels.push({ id, type: relation.type, target: `media/${target.split('/').at(-1)}` });
          embeddedCount++;
        } else if (relation.type.endsWith('/hyperlink') && relation.external) {
          rels.push({ ...relation, id });
        } else throw new Error(`${workbook.name} 的图片包含未嵌入内容，无法完整复制，请先将原图嵌入文件。`);
        relationMap.set(sourceId, id);
        node.setAttributeNS(NS.rel, attr.name, id);
      }
    }
    if (!embeddedCount) throw new Error(`${workbook.name} 的图片未包含可复制的原图，请重新保存原表。`);
    doc.documentElement.appendChild(item);
    copied.get(workbook).set(sourceName, name);
    return name;
  }

  function copyCell(workbook, cell, sourceName) {
    const newName = copy(workbook, sourceName);
    let formula = first(cell, 'f');
    const value = first(cell, 'v');
    if (!formula) {
      // A few WPS writers only save the cached DISPIMG expression.
      formula = el(cell.ownerDocument, 'f');
      formula.textContent = (value?.textContent || '').replace(/^=/, '');
      cell.insertBefore(formula, value);
    }
    if (newName !== sourceName) {
      const replaceName = text => text.replace(/((?:_xlfn\.)?DISPIMG\(\s*")([^"]+)(")/i, (match, prefix, id, suffix) => id === sourceName ? `${prefix}${newName}${suffix}` : match);
      formula.textContent = replaceName(formula.textContent);
      if (value) value.textContent = replaceName(value.textContent);
    }
    return newName;
  }

  return {
    copyCell,
    get count() { return references; },
    get objectCount() { return children(doc.documentElement).length; },
    files() { return { 'xl/cellimages.xml': xmlBytes(doc), 'xl/_rels/cellimages.xml.rels': relationships(rels) }; },
  };
}
