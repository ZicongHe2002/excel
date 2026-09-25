import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { strFromU8, strToU8 } from 'fflate';

export const NS = {
  sheet: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  package: 'http://schemas.openxmlformats.org/package/2006/relationships',
  drawing: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  art: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  types: 'http://schemas.openxmlformats.org/package/2006/content-types',
};
export function parseXml(data) {
  const text = typeof data === 'string' ? data : strFromU8(data);
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('不支持包含外部实体的 Excel 文件。');
  const errors = [];
  const doc = new DOMParser({ errorHandler: { warning() {}, error: e => errors.push(e), fatalError: e => errors.push(e) } }).parseFromString(text, 'application/xml');
  if (errors.length || !doc.documentElement) throw new Error('Excel 文件中的 XML 数据不完整。');
  return doc;
}
export const serialize = node => new XMLSerializer().serializeToString(node);
export const xmlBytes = node => strToU8(serialize(node));
export const children = (node, name) => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1 && (!name || n.localName === name));
export const all = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
export const first = (node, name) => children(node, name)[0] || null;
export const remove = node => node?.parentNode?.removeChild(node);
export function el(doc, name, attrs = {}, namespace = NS.sheet) {
  const node = doc.createElementNS(namespace, name);
  for (const [key, value] of Object.entries(attrs)) if (value != null) node.setAttribute(key, String(value));
  return node;
}
export function normalizePath(path) {
  const out = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!out.length) throw new Error('Excel 内部文件路径无效。'); out.pop(); }
    else out.push(part);
  }
  return out.join('/');
}
export const resolvePath = (base, target) => normalizePath(target.startsWith('/') ? target : `${base.slice(0, base.lastIndexOf('/') + 1)}${target}`);
export const relsPath = path => `${path.slice(0, path.lastIndexOf('/') + 1)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
export function readRels(files, path) {
  if (!files[relsPath(path)]) return [];
  return children(parseXml(files[relsPath(path)]).documentElement).map(n => ({
    id: n.getAttribute('Id'), type: n.getAttribute('Type'), target: n.getAttribute('Target'), external: n.getAttribute('TargetMode') === 'External',
    path: n.getAttribute('TargetMode') === 'External' ? null : resolvePath(path, n.getAttribute('Target')),
  }));
}
export function relationships(entries) {
  const doc = parseXml(`<Relationships xmlns="${NS.package}"/>`);
  for (const entry of entries) doc.documentElement.appendChild(el(doc, 'Relationship', { Id: entry.id, Type: entry.type, Target: entry.target, TargetMode: entry.external ? 'External' : null }, NS.package));
  return xmlBytes(doc);
}
export function cellRef(ref) {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref || '');
  if (!match) return null;
  return { col: match[1].toUpperCase(), row: Number(match[2]), column: columnNumber(match[1]) };
}
export const columnNumber = letters => [...letters.toUpperCase()].reduce((sum, c) => sum * 26 + c.charCodeAt(0) - 64, 0);
export function columnName(number) { let out = ''; while (number) { number--; out = String.fromCharCode(65 + number % 26) + out; number = Math.floor(number / 26); } return out; }
export function rangeRef(ref) { const [a, b = a] = ref.split(':'); return { start: cellRef(a), end: cellRef(b) }; }
export function mapRange(ref, rowMap) {
  const r = rangeRef(ref);
  if (!r.start || !r.end) return null;
  const selected = [...rowMap.keys()].filter(n => n >= r.start.row && n <= r.end.row);
  if (!selected.length) return null;
  const start = `${r.start.col}${rowMap.get(selected[0])}`;
  const end = `${r.end.col}${rowMap.get(selected.at(-1))}`;
  return start === end ? start : `${start}:${end}`;
}
