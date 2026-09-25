import { zipSync, strToU8 } from 'fflate';

export function makeFloatingImageWorkbook(imageBytes, { absolute = false, hyperlinks = false, svg = false } = {}) {
  const sheet = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const main = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const textCell = (ref, value) => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
  const marker = (tag, row, col) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
  const geometry = (id) => {
    const row = [1, 3, 4][id - 1];
    const column = id === 3 ? 3 : 2;
    const x = column * 140 * 9525;
    const y = 15 * 12700 + (row - 1) * 85 * 12700 + (absolute ? 95250 : 0);
    const width = id === 2 && !absolute ? 140 * 9525 : 952500;
    const height = id === 2 && !absolute ? 2 * 85 * 12700 : 952500;
    return `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`;
  };
  const pic = (id) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}">${hyperlinks ? '<a:hlinkClick r:id="rIdHyperlink"/>' : ''}</xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage">${svg ? '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst>' : ''}</a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${geometry(id)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
  const absoluteAnchor = (id, row, column) => `<xdr:absoluteAnchor><xdr:pos x="${column * 140 * 9525}" y="${15 * 12700 + (row - 1) * 85 * 12700 + 95250}"/><xdr:ext cx="952500" cy="952500"/>${pic(id)}<xdr:clientData/></xdr:absoluteAnchor>`;
  const anchors = absolute
    ? `${absoluteAnchor(1, 1, 2)}${absoluteAnchor(2, 3, 2)}${absoluteAnchor(3, 4, 3)}`
    : `<xdr:oneCellAnchor>${marker('from', 1, 2)}<xdr:ext cx="952500" cy="952500"/>${pic(1)}<xdr:clientData/></xdr:oneCellAnchor><xdr:twoCellAnchor>${marker('from', 3, 2)}${marker('to', 5, 3)}${pic(2)}<xdr:clientData/></xdr:twoCellAnchor><xdr:oneCellAnchor>${marker('from', 4, 3)}<xdr:ext cx="952500" cy="952500"/>${pic(3)}<xdr:clientData/></xdr:oneCellAnchor>`;
  const files = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="rIdWorkbook" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${sheet}" xmlns:r="${rel}"><sheets><sheet name="浮动图片" sheetId="1" r:id="rIdSheet"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="rIdSheet" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${sheet}" xmlns:r="${rel}"><dimension ref="A1:D5"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="4" width="20" customWidth="1"/></cols><sheetData><row r="1">${textCell('A1', '颜色')}${textCell('B1', '款号')}${textCell('C1', '图片')}${textCell('D1', '详细图片')}</row><row r="2" ht="85" customHeight="1">${textCell('A2', '米色')}${textCell('B2', 'FIRST')}</row><row r="3" ht="85" customHeight="1">${textCell('A3', '深灰')}</row><row r="4" ht="85" customHeight="1">${textCell('A4', '黑色')}${textCell('B4', 'SECOND')}</row><row r="5" ht="85" customHeight="1">${textCell('A5', '驼色')}</row></sheetData><mergeCells count="2"><mergeCell ref="B2:B3"/><mergeCell ref="B4:B5"/></mergeCells><drawing r:id="rIdDrawing"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="rIdDrawing" Type="${rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${drawing}" xmlns:a="${main}" xmlns:r="${rel}">${anchors}</xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': `<Relationships xmlns="${pkgRel}"><Relationship Id="rIdImage" Type="${rel}/image" Target="../media/image1.jpeg"/>${hyperlinks ? `<Relationship Id="rIdHyperlink" Type="${rel}/hyperlink" Target="https://example.com/products/second?color=black&amp;size=M" TargetMode="External"/>` : ''}${svg ? `<Relationship Id="rIdSvg" Type="${rel}/image" Target="../media/vector.svg"/>` : ''}</Relationships>`,
  };
  if (svg) {
    files['xl/media/vector.svg'] = '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="#115544"/></svg>';
    files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Default Extension="svg" ContentType="image/svg+xml"/></Types>');
  }
  return zipSync(Object.fromEntries([...Object.entries(files).map(([name, value]) => [name, strToU8(value)]), ['xl/media/image1.jpeg', imageBytes]]));
}
