import { NS, parseXml, children, all, first, el } from './xml.js';

const DEFAULT_STYLE = `<styleSheet xmlns="${NS.sheet}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

function themePalette(workbook) {
  const path = workbook.themePath || Object.keys(workbook.files).find(p => /theme\d*\.xml$/.test(p));
  if (!path || !workbook.files[path]) return [];
  const doc = parseXml(workbook.files[path]);
  const scheme = all(doc, 'clrScheme')[0];
  const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return names.map(name => {
    const node = children(scheme, name)[0];
    const color = children(node)[0];
    return color?.getAttribute('lastClr') || color?.getAttribute('val') || null;
  });
}

export function mergeStyles(workbooks) {
  const doc = parseXml(`<styleSheet xmlns="${NS.sheet}"/>`);
  const root = doc.documentElement;
  const names = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs'];
  const containers = Object.fromEntries(names.map(name => [name, root.appendChild(el(doc, name, { count: 0 }))]));
  const maps = new Map();
  let nextFormatId = 164;
  for (const [sourceIndex, workbook] of workbooks.entries()) {
    const path = workbook.stylesPath || Object.keys(workbook.files).find(p => /(^|\/)styles\.xml$/.test(p));
    const source = parseXml(path && workbook.files[path] ? workbook.files[path] : DEFAULT_STYLE);
    const palette = themePalette(workbook);
    const map = Object.fromEntries(names.map(name => [name, children(containers[name]).length]));
    const normal = children(first(source.documentElement, 'cellStyles')).find(n => n.getAttribute('builtinId') === '0');
    const normalXf = children(first(source.documentElement, 'cellStyleXfs'))[Number(normal?.getAttribute('xfId') || 0)];
    const font = children(first(source.documentElement, 'fonts'))[Number(normalXf?.getAttribute('fontId') || 0)];
    const fontName = first(font, 'name')?.getAttribute('val') || 'Calibri';
    const fontSize = Number(first(font, 'sz')?.getAttribute('val') || 11);
    // Reference metrics at 96 DPI. The actual cell anchor remains authoritative
    // when a spreadsheet viewer substitutes a missing font.
    const baseWidth = /微软雅黑|Microsoft YaHei/i.test(fontName) ? 9 : /Calibri|Carlito/i.test(fontName) ? 7 : 8;
    map.maximumDigitWidth = Math.max(1, Math.round(baseWidth * fontSize / 11));
    const formats = new Map();
    for (const fmt of children(first(source.documentElement, 'numFmts'))) {
      const clone = doc.importNode(fmt, true);
      const old = Number(fmt.getAttribute('numFmtId'));
      const id = nextFormatId++;
      formats.set(old, id);
      clone.setAttribute('numFmtId', id);
      containers.numFmts.appendChild(clone);
    }
    for (const name of names.slice(1)) {
      for (const original of children(first(source.documentElement, name))) {
        const clone = doc.importNode(original, true);
        for (const color of [clone, ...all(clone, '*')]) {
          if (color.hasAttribute('theme')) {
            const rgb = palette[Number(color.getAttribute('theme'))];
            if (rgb && /^[0-9a-f]{6}$/i.test(rgb)) {
              color.removeAttribute('theme'); color.setAttribute('rgb', `FF${rgb.toUpperCase()}`);
            }
          }
        }
        if (name === 'cellStyleXfs' || name === 'cellXfs') {
          for (const [attr, group] of [['fontId', 'fonts'], ['fillId', 'fills'], ['borderId', 'borders']]) {
            clone.setAttribute(attr, Number(clone.getAttribute(attr) || 0) + map[group]);
          }
          if (clone.hasAttribute('xfId')) clone.setAttribute('xfId', Number(clone.getAttribute('xfId')) + map.cellStyleXfs);
          if (formats.has(Number(clone.getAttribute('numFmtId')))) clone.setAttribute('numFmtId', formats.get(Number(clone.getAttribute('numFmtId'))));
        }
        if (name === 'cellStyles') {
          clone.setAttribute('xfId', Number(clone.getAttribute('xfId') || 0) + map.cellStyleXfs);
          if (sourceIndex) { clone.setAttribute('name', `${clone.getAttribute('name')}_${sourceIndex + 1}`); clone.removeAttribute('builtinId'); }
        }
        if (name === 'dxfs') for (const fmt of all(clone, 'numFmt')) {
          if (formats.has(Number(fmt.getAttribute('numFmtId')))) fmt.setAttribute('numFmtId', formats.get(Number(fmt.getAttribute('numFmtId'))));
        }
        containers[name].appendChild(clone);
      }
    }
    map.palette = palette;
    maps.set(workbook, map);
  }
  for (const container of Object.values(containers)) container.setAttribute('count', children(container).length);
  root.appendChild(el(doc, 'tableStyles', { count: 0, defaultTableStyle: 'TableStyleMedium2', defaultPivotStyle: 'PivotStyleLight16' }));
  return { doc, maps };
}
