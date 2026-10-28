import { copyFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';

const root = process.cwd();
const output = join(root, 'www');
await mkdir(join(output, 'assets'), { recursive: true });
for (const file of ['index.html', 'styles.css', 'app.js', 'data.js', 'excel.js', 'manifest.webmanifest', 'sw.js']) {
  await copyFile(join(root, file), join(output, file));
}
await copyFile(join(root, 'assets', 'icon.svg'), join(output, 'assets', 'icon.svg'));
console.log('移动端发布文件已生成：www/');
