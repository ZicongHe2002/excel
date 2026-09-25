import { build } from 'esbuild';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const result = await build({
  entryPoints: [path.join(root, 'src/main.js')], bundle: true, minify: true,
  format: 'iife', platform: 'browser', target: ['chrome110', 'safari16'],
  outfile: path.join(root, 'dist/app.js'), write: false, legalComments: 'inline',
});
const javascript = result.outputFiles.find(f => f.path.endsWith('.js')).text.replace(/<\/script/gi, '<\\/script');
const bundledCss = result.outputFiles.find(f => f.path.endsWith('.css'))?.text;
const css = bundledCss || await fs.readFile(path.join(root, 'src/style.css'), 'utf8');
let html = await fs.readFile(path.join(root, 'index.html'), 'utf8');
html = html.replace(/<script\b[^>]*\bsrc=["'][^"']*main\.js["'][^>]*>\s*<\/script>/gi, () => `<script>${javascript}</script>`);
html = html.replace(/<link\b[^>]*\bhref=["'][^"']*(?:style|main)\.css["'][^>]*>/gi, '');
html = html.replace('</head>', () => `<style>${css}</style></head>`);
if (/src=["'](?:\/|\.\/)?src\//.test(html)) throw new Error('页面仍引用源文件，无法独立运行。');
await fs.mkdir(path.join(root, 'dist'), { recursive: true });
await fs.writeFile(path.join(root, 'dist/index.html'), html);
await fs.writeFile(path.join(root, '选款助手.html'), html);
console.log(`已生成可双击打开的 选款助手.html（${Math.round(Buffer.byteLength(html) / 1024)} KB），无外部依赖。`);
