import { readFile } from 'node:fs/promises';

const requiredFiles = ['index.html','styles.css','app.js','data.js','manifest.webmanifest','sw.js','capacitor.config.json'];
for (const file of requiredFiles) await readFile(new URL(`../${file}`, import.meta.url));

const manifest = JSON.parse(await readFile(new URL('../manifest.webmanifest', import.meta.url), 'utf8'));
const capacitor = JSON.parse(await readFile(new URL('../capacitor.config.json', import.meta.url), 'utf8'));
if (manifest.display !== 'standalone') throw new Error('PWA manifest must use standalone mode');
if (capacitor.appId !== 'com.qizhang.offline') throw new Error('Unexpected Android package id');
console.log('静态检查通过：核心文件、PWA 配置和 Android 配置均有效。');
