// 把 charts.html 裡的每一張圖表輸出成 3 倍解析度 PNG
import { chromium } from 'playwright-core';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const exe = process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';

const b = await chromium.launch({ executablePath: exe, args: ['--no-sandbox'] });
const p = await b.newPage({ viewport: { width: 1400, height: 900 }, deviceScaleFactor: 3 });
p.on('pageerror', e => console.log('ERR', e.message));
await p.goto('file://' + join(here, 'charts.html'));
const ids = await p.$$eval('.c', e => e.map(x => x.id));
for (const id of ids) await (await p.$('#' + id)).screenshot({ path: join(here, id + '.png') });
console.log('rendered:', ids.join(', '));
await b.close();
