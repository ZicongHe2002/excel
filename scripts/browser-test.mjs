import { chromium } from 'playwright';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import http from 'node:http';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { readWorkbook } from '../src/importer.js';
import { inspectXlsx, assertValidPackage } from '../tests/xlsx-inspector.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'test-results');
await fs.mkdir(output, { recursive: true });
const renderMode = process.argv.includes('--render');
let browser;
let server;
let pageUrl = pathToFileURL(path.join(root, '选款助手.html')).href;
const errors = [];
const remoteRequests = [];
try {
  if (renderMode) {
    const html = await fs.readFile(path.join(root, 'Render上传包/public/index.html'));
    server = http.createServer((request, response) => {
      if (request.url !== '/' && request.url !== '/index.html') {
        response.writeHead(404).end();
        return;
      }
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }).end(html);
    });
    await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
    pageUrl = `http://127.0.0.1:${server.address().port}/`;
  }
  browser = await chromium.launch({ executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: true });
  const context = await browser.newContext({ viewport: { width: 1440, height: 1050 }, acceptDownloads: true, offline: !renderMode, permissions: ['clipboard-read', 'clipboard-write'] });
  const page = await context.newPage();
  page.on('pageerror', error => errors.push(error.message));
  page.on('request', request => {
    if (/^https?:/.test(request.url()) && (!renderMode || new URL(request.url()).origin !== new URL(pageUrl).origin)) remoteRequests.push(request.url());
  });
  await page.goto(pageUrl);
  await page.getByRole('heading', { name: '商品资料', exact: true }).waitFor();
  assert.equal(await page.locator('#export-button').isDisabled(), true);
  assert.equal(await page.locator('#sku-search').isDisabled(), true);
  assert.equal(await page.locator('#copy-codes').isDisabled(), true);
  await page.screenshot({ path: path.join(output, 'desktop-empty.png'), fullPage: true });
  await page.getByRole('button', { name: '使用说明', exact: true }).click();
  assert.equal(await page.locator('#help-dialog').isVisible(), true);
  await page.keyboard.press('Escape');
  assert.equal(await page.locator('#help-dialog').isVisible(), false);
  await page.locator('#file-input').setInputFiles([path.join(root, 'KOKO选款--8.5信息表(1).xlsx'), path.join(root, 'Cole Haan-9.4选款.xlsx')]);
  await page.waitForFunction(() => document.getElementById('file-count').textContent === '2 份文件' && !document.getElementById('file-input').disabled);

  // Browse and find complete SKU blocks without introducing a colour-selection step.
  const search = page.locator('#sku-search');
  const skuResults = page.locator('#sku-results .sku-result');
  assert.equal(await search.isDisabled(), false);
  assert.equal(await skuResults.count(), 24, 'the two source catalogues combine into 24 distinct SKUs');
  await search.fill('h140');
  assert.equal(await skuResults.count(), 1);
  assert.equal(await skuResults.locator('.sku-result-code').innerText(), 'H140');
  assert.match(await skuResults.locator('.sku-result-meta').innerText(), /2.*来源.*11\s*行/);
  const sourceNames = await skuResults.locator('.sku-result-sources').innerText();
  assert.ok(sourceNames.includes('KOKO选款--8.5信息表(1).xlsx'));
  assert.ok(sourceNames.includes('Cole Haan-9.4选款.xlsx'));
  await skuResults.locator('.sku-copy-button').click();
  await page.waitForFunction(() => /已复制货号 H140/.test(document.getElementById('search-status').textContent));
  assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'H140', 'individual copying writes the visible SKU');
  await skuResults.locator('.sku-add-button').click();
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140');
  assert.equal(await skuResults.locator('.sku-add-button').isDisabled(), true);
  assert.equal(await skuResults.locator('.sku-add-button').innerText(), '已加入');
  assert.match(await page.locator('#export-caption').innerText(), /2 个商品区块.*11 行/);
  await search.press('Enter');
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140', 'adding an already selected code never duplicates it');
  await search.press('Escape');
  assert.equal(await search.inputValue(), '');
  assert.equal(await skuResults.count(), 24);

  await search.fill('lr');
  const partialCodes = await skuResults.locator('.sku-result-code').allTextContents();
  assert.ok(partialCodes.length > 1 && partialCodes.every(code => code.includes('LR')), 'search is case-insensitive and matches part of a SKU');
  await search.press('Enter');
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140', 'an ambiguous search must not silently add a SKU');
  assert.match(await page.locator('#search-status').innerText(), /多个货号/);
  await search.fill('lr05');
  assert.equal(await skuResults.count(), 1);
  await search.press('Enter');
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140\nLR059', 'Enter adds the sole partial-search result');
  await search.fill('t026');
  await search.press('Enter');
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140\nLR059\nT026', 'Enter also adds an exact match');
  await search.fill('lr059');
  assert.equal(await skuResults.locator('.sku-add-button').isDisabled(), true);
  await page.locator('#sku-input').fill('H140\nT026');
  assert.equal(await skuResults.locator('.sku-add-button').isDisabled(), false, 'editing the customer list updates search actions');
  await skuResults.locator('.sku-add-button').click();
  assert.equal(await page.locator('#sku-input').inputValue(), 'H140\nT026\nLR059');
  await page.locator('#sku-input').fill('h140，lr059\nH140;t026');
  await page.locator('#copy-codes').click();
  await page.waitForFunction(() => /已复制清单/.test(document.getElementById('copy-status').textContent));
  assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'H140\nLR059\nT026', 'copying the list normalizes, deduplicates and preserves input order');

  // Local-file browsers may reject the modern clipboard API. Exercise the real
  // textarea-copy fallback, then reject both methods and check the honest failure message.
  await search.fill('h140');
  await page.evaluate(() => {
    const test = window.__copyFallbackTest = {
      writeDescriptor: Object.getOwnPropertyDescriptor(navigator.clipboard, 'writeText'),
      execDescriptor: Object.getOwnPropertyDescriptor(document, 'execCommand'),
      originalExec: document.execCommand,
      count: 0,
      payload: '',
      forceFailure: false,
    };
    Object.defineProperty(navigator.clipboard, 'writeText', { configurable: true, value: async () => { throw new DOMException('Test clipboard denial', 'NotAllowedError'); } });
    Object.defineProperty(document, 'execCommand', { configurable: true, value(command, ...args) {
      if (command === 'copy') {
        test.count++;
        const input = document.activeElement;
        test.payload = input?.value?.slice(input.selectionStart, input.selectionEnd) || '';
        if (test.forceFailure) return false;
      }
      return test.originalExec.call(document, command, ...args);
    } });
  });
  try {
    await skuResults.locator('.sku-copy-button').click();
    await page.waitForFunction(() => /已复制货号 H140|无法自动复制/.test(document.getElementById('search-status').textContent));
    const fallback = await page.evaluate(() => ({ count: window.__copyFallbackTest.count, payload: window.__copyFallbackTest.payload }));
    assert.equal(fallback.count, 1, 'clipboard denial invokes the textarea copy fallback');
    assert.equal(fallback.payload, 'H140', 'the fallback selects exactly the requested SKU');
    if (/已复制/.test(await page.locator('#search-status').innerText())) {
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'H140', 'the real fallback writes the requested SKU');
    } else {
      assert.match(await page.locator('#search-status').innerText(), /Ctrl\+C/);
    }
    await page.evaluate(() => { window.__copyFallbackTest.forceFailure = true; document.getElementById('search-status').textContent = ''; });
    await skuResults.locator('.sku-copy-button').click();
    await page.waitForFunction(() => /无法自动复制/.test(document.getElementById('search-status').textContent));
    assert.doesNotMatch(await page.locator('#search-status').innerText(), /已复制/);
    assert.match(await page.locator('#search-status').innerText(), /Ctrl\+C/);
    assert.equal(await skuResults.locator('.sku-result-code').innerText(), 'H140', 'the code remains available to select manually after copying fails');
    assert.equal(await page.locator('.clipboard-buffer').count(), 0, 'temporary copy textareas are removed');
  } finally {
    await page.evaluate(() => {
      const test = window.__copyFallbackTest;
      if (test.writeDescriptor) Object.defineProperty(navigator.clipboard, 'writeText', test.writeDescriptor);
      else delete navigator.clipboard.writeText;
      if (test.execDescriptor) Object.defineProperty(document, 'execCommand', test.execDescriptor);
      else delete document.execCommand;
      delete window.__copyFallbackTest;
    });
  }
  await search.fill('NOT_FOUND');
  assert.equal(await skuResults.count(), 0);
  assert.match(await page.locator('#sku-results').innerText(), /没有找到/);
  await search.press('Enter');
  assert.match(await page.locator('#search-status').innerText(), /没有找到/);
  await search.press('Escape');
  await page.locator('#clear-codes').click();
  assert.equal(await page.locator('#sku-input').inputValue(), '');
  assert.equal(await page.locator('#copy-codes').isDisabled(), true);
  assert.equal(await page.locator('#sku-results .sku-add-button:disabled').count(), 0, 'clearing selection re-enables every SKU');
  await page.screenshot({ path: path.join(output, 'desktop-search.png'), fullPage: true });

  for (const code of ['lr059', 'H140', 'T026']) {
    await search.fill(code);
    await skuResults.locator('.sku-add-button').click();
  }
  await search.press('Escape');
  await page.locator('#sku-input').fill(`${await page.locator('#sku-input').inputValue()}\nLR059\nNOT_FOUND`);
  assert.equal(await page.locator('#matched-count').innerText(), '3');
  assert.equal(await page.locator('#missing-count').innerText(), '1');
  await page.locator('#filename-input').fill('客户选款验证');
  await page.screenshot({ path: path.join(output, 'desktop-ready.png'), fullPage: true });
  const downloadEvent = page.waitForEvent('download');
  await page.locator('#export-button').click();
  const download = await downloadEvent;
  assert.equal(download.suggestedFilename(), '客户选款验证.xlsx');
  await download.saveAs(path.join(output, '客户选款验证.xlsx'));
  await page.locator('#result-panel').waitFor({ state: 'visible' });
  assert.match(await page.locator('#result-description').innerText(), /14 行资料/);
  const exported = readWorkbook(new Uint8Array(await fs.readFile(path.join(output, '客户选款验证.xlsx'))), { name: '客户选款验证.xlsx' });
  assert.equal(exported.sheets.length, 2);
  assert.equal(exported.rowCount, 14);
  const pictures = inspectXlsx(new Uint8Array(await fs.readFile(path.join(output, '客户选款验证.xlsx'))));
  assertValidPackage(pictures);
  assert.equal(pictures.cellImages.size, 17, 'all 17 selected original cell image objects are embedded');
  assert.equal(pictures.sheets.reduce((sum, sheet) => sum + sheet.drawings.length, 0), 0, 'original cell images are not replaced with floating pictures');
  assert.match(await page.locator('#result-description').innerText(), /17 张单元格图片.*WPS/);
  await page.screenshot({ path: path.join(output, 'desktop-result.png'), fullPage: true });
  await page.locator('#sku-input').fill('NO_SUCH_SKU');
  assert.equal(await page.locator('#export-button').isDisabled(), true);
  await page.locator('#file-input').setInputFiles({ name: '损坏表格.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('invalid') });
  await page.waitForFunction(() => !document.getElementById('file-input').disabled);
  assert.match(await page.locator('#import-notes').innerText(), /有效|损坏/);
  assert.equal(await page.locator('#file-count').innerText(), '2 份文件');
  await page.locator('#sku-input').fill('LR059');
  await page.setViewportSize({ width: 390, height: 844 });
  await page.screenshot({ path: path.join(output, 'mobile.png'), fullPage: true });
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth), false, 'mobile page should not overflow');
  await page.locator('#clear-files').click();
  assert.equal(await page.locator('#file-count').innerText(), '0 份文件');
  assert.equal(await page.locator('#export-button').isDisabled(), true);
  assert.equal(await page.locator('#sku-search').isDisabled(), true);
  assert.equal(await skuResults.count(), 0, 'clearing source files clears the search catalogue');
  assert.deepEqual(errors, [], 'no JavaScript errors');
  assert.deepEqual(remoteRequests, [], 'standalone app never makes a network request');
  console.log(`浏览器端验证通过：${renderMode ? 'Render 上传包通过 HTTP 打开' : '离线打开'}、多文件上传、货号检索、逐项加入、复制货号/清单、剪贴板拒绝降级、重复输入、同号多来源、下载、错误文件、清空、手机布局；无脚本报错和外部网络请求。`);
} finally {
  await browser?.close();
  if (server) await new Promise(resolve => server.close(resolve));
}
