import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

// Build the same self-contained page used by the offline version.
await import('./build.mjs');

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packageDir = path.join(root, 'Render上传包');
const page = await fs.readFile(path.join(root, 'dist/index.html'));
const guide = await fs.readFile(path.join(root, 'Render上传说明.md'));
const config = strToU8(`services:
  - type: web
    name: sku-selection
    runtime: static
    buildCommand: test -s public/index.html
    staticPublishPath: public
    autoDeployTrigger: commit
`);

// Explicitly allow only deployment artifacts; never collect the workspace's
// source Excel files, product images, test outputs or installed dependencies.
const files = {
  'public/index.html': page,
  'render.yaml': config,
  'README.md': guide,
};
const allowed = new Set([...Object.keys(files), 'public']);
async function rejectUnexpectedFiles(directory, prefix = '') {
  for (const entry of await fs.readdir(directory, { withFileTypes: true }).catch((error) => {
    if (error.code === 'ENOENT') return [];
    throw error;
  })) {
    const relative = prefix + entry.name;
    if (!allowed.has(relative) || entry.isSymbolicLink()) {
      throw new Error(`上传包中有额外文件，请先移走后重新打包：${relative}`);
    }
    if (entry.isDirectory()) await rejectUnexpectedFiles(path.join(directory, entry.name), `${relative}/`);
  }
}
await rejectUnexpectedFiles(packageDir);
for (const [relative, contents] of Object.entries(files)) {
  const target = path.join(packageDir, relative);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, contents);
}
const zipPath = path.join(root, '选款助手-Render上传包.zip');
await fs.writeFile(zipPath, zipSync(files, { level: 9 }));
console.log('已生成 Render上传包/ 和 选款助手-Render上传包.zip；包含网页、部署配置和说明，共 3 个文件。');
