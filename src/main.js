import './style.css';
import { readWorkbook, parseSkuInput, matchSelection, normalizeSku } from './importer.js';
import { exportSelection } from './exporter.js';

const $ = (id) => document.getElementById(id);
const state = { workbooks: [], catalog: [], searchLimit: 50, importing: false, exporting: false, notes: [], importSequence: 0, signatures: new Set() };
const fileInput = $('file-input');
const skuInput = $('sku-input');
const skuSearch = $('sku-search');
const dropZone = $('drop-zone');
const exportButton = $('export-button');
const emptyFileMarkup = $('file-list').innerHTML;
const excelIcon = '<svg viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M14 3H6a1.5 1.5 0 0 0-1.5 1.5v15A1.5 1.5 0 0 0 6 21h12a1.5 1.5 0 0 0 1.5-1.5V8.5L14 3Z" stroke="currentColor" stroke-width="1.3"/><path d="M14 3v5.5h5.5M8 12l6 5m0-5-6 5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const deleteIcon = '<svg viewBox="0 0 20 20" fill="none" aria-hidden="true"><path d="M6 6l8 8M14 6l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';

function element(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined) el.textContent = text;
  return el;
}

function warningText(warning) {
  return typeof warning === 'string' ? warning : warning?.message || warning?.text || '部分内容需要留意，请核对导出的文件。';
}

function errorText(error) {
  return error?.message || '文件处理失败，请确认文件是有效的 .xlsx 商品表后重试。';
}

function parseInput() {
  const parsed = parseSkuInput(skuInput.value);
  return { codes: parsed.codes || [], duplicates: Array.isArray(parsed.duplicates) ? parsed.duplicates.length : Number(parsed.duplicates) || 0 };
}

function getMatch(codes) {
  return matchSelection(state.workbooks, codes);
}

function formatCount(count) {
  return Number(count || 0).toLocaleString('zh-CN');
}

function hideResult() {
  $('result-panel').hidden = true;
}

function buildCatalog() {
  const catalog = new Map();
  for (const workbook of state.workbooks) {
    for (const sheet of workbook.sheets) {
      for (const block of sheet.blocks) {
        const code = block.normalizedSku;
        if (!catalog.has(code)) catalog.set(code, { code, rowCount: 0, sources: new Map() });
        const entry = catalog.get(code);
        entry.rowCount += block.end - block.start + 1;
        entry.sources.set(`${workbook.id}\u0000${sheet.path}`, workbook.sheets.length > 1 ? `${workbook.name} / ${sheet.name}` : workbook.name);
      }
    }
  }
  return [...catalog.values()];
}

function searchCatalog() {
  const query = normalizeSku(skuSearch.value);
  const matches = state.catalog.filter((entry) => entry.code.includes(query));
  // Keep an exact match first, then retain the order in the original files.
  return [...matches.filter((entry) => entry.code === query), ...matches.filter((entry) => entry.code !== query)];
}

function renderSearch() {
  const busy = state.importing || state.exporting;
  const selected = new Set(parseInput().codes);
  const matches = searchCatalog();
  skuSearch.disabled = busy || !state.catalog.length;
  $('catalog-count').textContent = skuSearch.value.trim()
    ? `${formatCount(matches.length)} / ${formatCount(state.catalog.length)} 个货号`
    : `${formatCount(state.catalog.length)} 个货号`;
  const list = $('sku-results');
  const scrollTop = list.scrollTop;
  list.replaceChildren();
  if (!matches.length) {
    list.append(element('p', 'sku-search-empty', state.catalog.length ? '没有找到这个货号，试试只输入其中几个字符。' : '添加商品资料后，就能在这里搜索和复制货号。'));
  }
  for (const entry of matches.slice(0, state.searchLimit)) {
    const item = element('div', 'sku-result');
    item.setAttribute('role', 'listitem');
    const info = element('div', 'sku-result-info');
    const sources = [...entry.sources.values()].join('；');
    const sourceText = element('div', 'sku-result-sources', sources);
    sourceText.title = sources;
    info.append(
      element('strong', 'sku-result-code', entry.code),
      element('span', 'sku-result-meta', `${entry.sources.size} 个来源 · ${formatCount(entry.rowCount)} 行资料`),
      sourceText,
    );
    const actions = element('div', 'sku-result-actions');
    const copy = element('button', 'sku-copy-button', '复制货号');
    copy.type = 'button';
    copy.disabled = busy;
    copy.setAttribute('aria-label', `复制货号 ${entry.code}`);
    copy.addEventListener('click', () => copyText(entry.code, `已复制货号 ${entry.code}`));
    const add = element('button', 'sku-add-button', selected.has(entry.code) ? '已加入' : '加入清单');
    add.type = 'button';
    add.disabled = busy || selected.has(entry.code);
    add.setAttribute('aria-label', `${selected.has(entry.code) ? '已加入' : '加入清单'} ${entry.code}`);
    add.addEventListener('click', () => addSku(entry));
    actions.append(copy, add);
    item.append(info, actions);
    list.append(item);
  }
  list.scrollTop = scrollTop;
  const more = $('search-more');
  more.hidden = matches.length <= state.searchLimit;
  more.disabled = busy;
  more.textContent = `显示更多（还有 ${formatCount(Math.max(0, matches.length - state.searchLimit))} 个）`;
}

function addSku(entry) {
  if (state.importing || state.exporting) return;
  if (parseInput().codes.includes(entry.code)) {
    $('search-status').textContent = `${entry.code} 已在客户清单中，无需重复加入。`;
    return;
  }
  skuInput.value = `${skuInput.value.trim()}${skuInput.value.trim() ? '\n' : ''}${entry.code}`;
  $('copy-status').textContent = '';
  hideResult();
  renderMatch();
  $('search-status').textContent = `已加入 ${entry.code}，将完整提取 ${formatCount(entry.rowCount)} 行资料。`;
  skuSearch.focus();
  skuSearch.select();
}

async function copyText(text, successMessage, statusId = 'search-status') {
  let copied = false;
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      copied = true;
    }
  } catch { /* Local HTML pages can deny the modern clipboard API. */ }
  if (!copied) {
    const active = document.activeElement;
    const inputSelection = active && typeof active.selectionStart === 'number'
      ? [active.selectionStart, active.selectionEnd, active.selectionDirection] : null;
    const selection = window.getSelection();
    const ranges = selection ? Array.from({ length: selection.rangeCount }, (_, i) => selection.getRangeAt(i).cloneRange()) : [];
    const temporary = element('textarea', 'clipboard-buffer');
    temporary.value = text;
    temporary.setAttribute('readonly', '');
    document.body.append(temporary);
    temporary.select();
    try { copied = document.execCommand('copy'); } catch { /* The user can still select and copy the visible text. */ }
    temporary.remove();
    if (active?.isConnected) active.focus({ preventScroll: true });
    if (inputSelection) active.setSelectionRange(...inputSelection);
    else if (selection) {
      selection.removeAllRanges();
      for (const range of ranges) selection.addRange(range);
    }
  }
  $(statusId).textContent = copied ? successMessage : '无法自动复制，请选中货号文字或客户清单，按 Ctrl+C（Mac 按 ⌘C）复制。';
}

function renderFiles() {
  const list = $('file-list');
  list.replaceChildren();
  if (!state.workbooks.length) list.innerHTML = emptyFileMarkup;
  for (const workbook of state.workbooks) {
    const item = element('div', 'file-item');
    const icon = element('span', 'file-icon');
    icon.innerHTML = excelIcon;
    const info = element('div', 'file-info');
    const name = element('div', 'file-name', workbook.name);
    name.title = workbook.name;
    const meta = element('div', 'file-meta');
    meta.append(element('span', '', `${formatCount(workbook.skuCount)} 个货号`), element('i'), element('span', '', `${formatCount(workbook.rowCount)} 行商品资料`));
    info.append(name, meta);
    const remove = element('button', 'icon-button');
    remove.type = 'button';
    remove.setAttribute('aria-label', `移除 ${workbook.name}`);
    remove.title = '移除文件';
    remove.innerHTML = deleteIcon;
    remove.disabled = state.importing || state.exporting;
    remove.addEventListener('click', () => {
      state.workbooks = state.workbooks.filter((book) => book.id !== workbook.id);
      state.signatures.delete(workbook._uploadSignature);
      state.notes = state.notes.filter((note) => note.workbookId !== workbook.id);
      hideResult();
      render();
    });
    item.append(icon, info, remove);
    list.append(item);
  }
  $('file-count').textContent = `${state.workbooks.length} 份文件`;
}

function renderNotes() {
  const container = $('import-notes');
  container.replaceChildren();
  if (state.importing) container.append(element('p', 'import-note info', '正在读取商品资料，请稍候…'));
  for (const note of state.notes) container.append(element('p', `import-note ${note.type || ''}`, note.text));
  const warnings = state.workbooks.flatMap((book) => (book.warnings || []).map((warning) => `${book.name}：${warningText(warning)}`));
  if (warnings.length) {
    const details = element('details', 'import-warning-details');
    details.append(element('summary', '', `${warnings.length} 条资料提示`));
    const list = element('ul');
    for (const warning of warnings) list.append(element('li', '', warning));
    details.append(list);
    container.append(details);
  }
}

function renderMatch() {
  const { codes, duplicates } = parseInput();
  const match = getMatch(codes);
  const matched = match.matchedCodes || [];
  const missing = match.missingCodes || [];
  const canMatch = state.workbooks.length > 0;
  $('sku-count').textContent = `${formatCount(codes.length)} 个货号`;
  $('matched-count').textContent = formatCount(matched.length);
  $('missing-count').textContent = canMatch ? formatCount(missing.length) : '0';
  $('duplicate-note').textContent = duplicates ? `已合并 ${duplicates} 个重复货号` : '';
  const missingNotice = $('missing-notice');
  missingNotice.hidden = !canMatch || !missing.length;
  if (!missingNotice.hidden) missingNotice.textContent = `未找到：${missing.join('、')}。${matched.length ? '导出时将包含其余已匹配货号。' : '请检查货号，或添加包含这些货号的商品资料。'}`;
  const busy = state.importing || state.exporting;
  exportButton.disabled = busy || !canMatch || !codes.length || !matched.length;
  exportButton.classList.toggle('is-loading', state.exporting);
  exportButton.setAttribute('aria-busy', String(state.exporting));
  $('export-label').textContent = state.exporting ? '正在生成选款表…' : '生成并下载选款表';
  $('clear-codes').disabled = busy || !skuInput.value;
  $('copy-codes').disabled = busy || !codes.length;
  $('clear-files').disabled = busy || !state.workbooks.length;
  dropZone.disabled = busy;
  fileInput.disabled = busy;
  skuInput.disabled = state.exporting;
  $('filename-input').disabled = state.exporting;
  let caption = '添加商品资料并输入货号，即可开始选款';
  if (state.exporting) caption = '正在保留原表图片与格式，请稍候';
  else if (state.importing) caption = '正在读取 Excel 文件';
  else if (canMatch && !codes.length) caption = '商品资料已就绪，粘贴客户货号即可导出';
  else if (canMatch && codes.length && !matched.length) caption = '暂未找到匹配货号，请核对输入或添加资料';
  else if (matched.length) caption = `将提取 ${formatCount(match.blockCount)} 个商品区块，共 ${formatCount(match.rowCount)} 行完整资料`;
  $('export-caption').textContent = caption;
  renderSearch();
}

function render() {
  renderFiles();
  renderNotes();
  state.catalog = buildCatalog();
  renderMatch();
}

async function importFiles(files) {
  if (state.importing || state.exporting || !files.length) return;
  state.importing = true;
  state.notes = [];
  hideResult();
  render();
  // Let the browser paint the reading state before parsing large files.
  await new Promise((resolve) => setTimeout(resolve, 30));
  for (const file of files) {
    if (!/\.xlsx$/i.test(file.name)) {
      state.notes.push({ type: 'error', text: `${file.name}：请使用 .xlsx 文件；旧版 .xls 可先在 Excel 或 WPS 中另存为 .xlsx。` });
      continue;
    }
    const signature = `${file.name}\u0000${file.size}\u0000${file.lastModified}`;
    if (state.signatures.has(signature)) {
      state.notes.push({ type: 'info', text: `${file.name} 已添加，本次已跳过。` });
      continue;
    }
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const id = `workbook-${Date.now()}-${++state.importSequence}`;
      const workbook = await readWorkbook(bytes, { name: file.name, id });
      workbook.id = workbook.id || id;
      workbook.name = workbook.name || file.name;
      workbook._uploadSignature = signature;
      state.workbooks.push(workbook);
      state.signatures.add(signature);
    } catch (error) {
      state.notes.push({ type: 'error', text: `${file.name}：${errorText(error)}` });
    }
    render();
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  state.importing = false;
  fileInput.value = '';
  render();
}

function defaultFilename() {
  const now = new Date();
  const date = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  return `选款表_${date}`;
}

function filename() {
  const input = $('filename-input').value.trim().replace(/\.xlsx$/i, '');
  return `${(input || defaultFilename()).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/, '') || defaultFilename()}.xlsx`;
}

function showResult({ title, description, warnings = [], error = false }) {
  const panel = $('result-panel');
  panel.hidden = false;
  panel.classList.toggle('error', error);
  $('result-icon').textContent = error ? '!' : '✓';
  $('result-title').textContent = title;
  $('result-description').textContent = description;
  const warningsContainer = $('result-warnings');
  warningsContainer.replaceChildren();
  if (warnings.length) {
    const list = element('ul');
    for (const warning of warnings) list.append(element('li', '', warningText(warning)));
    warningsContainer.append(list);
  }
}

async function exportWorkbook() {
  if (exportButton.disabled) return;
  state.exporting = true;
  hideResult();
  render();
  await new Promise((resolve) => setTimeout(resolve, 30));
  try {
    const { codes } = parseInput();
    const result = await exportSelection(state.workbooks, codes);
    const outputFilename = filename();
    const blob = new Blob([result.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = outputFilename;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60_000);
    const missing = result.missingCodes || [];
    const pictureNote = result.nativeImageCount ? ` ${formatCount(result.nativeImageCount)} 张单元格图片已按原表形式复制，请用 WPS 打开这份图片表。` : '';
    const summary = `已生成 ${outputFilename}，包含 ${(result.matchedCodes || []).length} 个货号、${formatCount(result.blockCount)} 个商品区块、${formatCount(result.rowCount)} 行资料，分为 ${formatCount(result.sheetCount)} 个工作表。${pictureNote}`;
    showResult({
      title: missing.length ? `选款表已生成 · ${missing.length} 个货号未找到` : '选款表已生成，下载已开始',
      description: summary,
      warnings: [...(missing.length ? [`未找到的货号未包含在文件中：${missing.join('、')}`] : []), ...(result.warnings || [])],
    });
  } catch (error) {
    showResult({ title: '选款表未能生成', description: errorText(error), error: true });
  } finally {
    state.exporting = false;
    render();
  }
}

fileInput.addEventListener('change', () => importFiles(Array.from(fileInput.files || [])));
dropZone.addEventListener('click', () => fileInput.click());
let dragDepth = 0;
dropZone.addEventListener('dragenter', (event) => { event.preventDefault(); if (!dropZone.disabled) { dragDepth++; dropZone.classList.add('dragging'); } });
dropZone.addEventListener('dragover', (event) => { event.preventDefault(); if (event.dataTransfer) event.dataTransfer.dropEffect = dropZone.disabled ? 'none' : 'copy'; });
dropZone.addEventListener('dragleave', (event) => { event.preventDefault(); dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) dropZone.classList.remove('dragging'); });
dropZone.addEventListener('drop', (event) => { event.preventDefault(); dragDepth = 0; dropZone.classList.remove('dragging'); importFiles(Array.from(event.dataTransfer?.files || [])); });
// Dropping outside the target should never navigate away and lose the selection.
window.addEventListener('dragover', (event) => { if (event.dataTransfer?.types.includes('Files')) event.preventDefault(); });
window.addEventListener('drop', (event) => { if (event.dataTransfer?.types.includes('Files')) event.preventDefault(); });
skuInput.addEventListener('input', () => { hideResult(); $('search-status').textContent = ''; $('copy-status').textContent = ''; renderMatch(); });
skuSearch.addEventListener('input', () => {
  state.searchLimit = 50;
  $('search-status').textContent = '';
  $('sku-results').scrollTop = 0;
  renderSearch();
});
skuSearch.addEventListener('keydown', (event) => {
  if (event.isComposing) return;
  if (event.key === 'Escape') {
    event.preventDefault();
    skuSearch.value = '';
    skuSearch.dispatchEvent(new Event('input'));
  } else if (event.key === 'Enter') {
    event.preventDefault();
    const matches = searchCatalog();
    const exact = matches.find((entry) => entry.code === normalizeSku(skuSearch.value));
    if (exact || matches.length === 1) addSku(exact || matches[0]);
    else $('search-status').textContent = matches.length ? '找到多个货号，请点击对应货号的“加入清单”。' : '没有找到这个货号，请核对输入。';
  }
});
$('search-more').addEventListener('click', () => { state.searchLimit += 50; renderSearch(); });
$('copy-codes').addEventListener('click', () => {
  const { codes } = parseInput();
  if (codes.length) copyText(codes.join('\n'), `已复制清单，共 ${formatCount(codes.length)} 个货号（已去重）。`, 'copy-status');
});
$('clear-codes').addEventListener('click', () => { skuInput.value = ''; hideResult(); $('search-status').textContent = ''; $('copy-status').textContent = ''; renderMatch(); skuInput.focus(); });
$('clear-files').addEventListener('click', () => { state.workbooks = []; state.signatures.clear(); state.notes = []; skuSearch.value = ''; state.searchLimit = 50; $('search-status').textContent = ''; hideResult(); render(); });
exportButton.addEventListener('click', exportWorkbook);
$('dismiss-result').addEventListener('click', hideResult);
const helpDialog = $('help-dialog');
$('help-button').addEventListener('click', () => helpDialog.showModal());
$('close-help').addEventListener('click', () => helpDialog.close());
$('help-done').addEventListener('click', () => { helpDialog.close(); dropZone.focus(); });
helpDialog.addEventListener('click', (event) => { if (event.target === helpDialog) { const rect = helpDialog.getBoundingClientRect(); if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) helpDialog.close(); } });
$('filename-input').placeholder = defaultFilename();
render();
