// Рендерит src/icon.svg в PNG (192/512) и вписывает их как base64 во все места
// index.html, где иконка встроена инлайново: favicon, apple-touch-icon (192/512),
// и manifest (any 192, any 512, maskable 512). Совпадение ищется по структурным
// якорям (rel=... / "sizes":...), а не по содержимому предыдущего base64 —
// поэтому работает одинаково при любой предыдущей иконке.
import fs from 'node:fs';
import sharp from 'sharp';

const svg = fs.readFileSync('src/icon.svg');

async function renderPngBase64(size) {
  const buf = await sharp(svg).resize(size, size).png().toBuffer();
  fs.writeFileSync(`src/icon-${size}.png`, buf);
  return buf.toString('base64');
}

const png192 = await renderPngBase64(192);
const png512 = await renderPngBase64(512);

let html = fs.readFileSync('index.html', 'utf8');
let replacements = 0;

const replaceOne = (re, b64) => {
  const before = html;
  html = html.replace(re, (_, pre, post) => { replacements++; return pre + b64 + post; });
  if (html === before) throw new Error(`Pattern not found: ${re}`);
};
const replaceAll = (re, b64) => {
  const before = html;
  html = html.replace(re, (_, pre, post) => { replacements++; return pre + b64 + post; });
  if (html === before) throw new Error(`Pattern not found: ${re}`);
};

// <link rel="icon" ... sizes="192x192" href="data:image/png;base64,...">
replaceOne(/(<link rel="icon"[^>]*sizes="192x192"[^>]*href="data:image\/png;base64,)[^"]+(")/, png192);
// <link rel="apple-touch-icon" sizes="192x192" href="data:image/png;base64,...">
replaceOne(/(<link rel="apple-touch-icon"[^>]*sizes="192x192"[^>]*href="data:image\/png;base64,)[^"]+(")/, png192);
// <link rel="apple-touch-icon" sizes="512x512" href="data:image/png;base64,...">
replaceOne(/(<link rel="apple-touch-icon"[^>]*sizes="512x512"[^>]*href="data:image\/png;base64,)[^"]+(")/, png512);
// manifest icons array: {"src":"data:image/png;base64,...","sizes":"192x192",...}
replaceOne(/("src":"data:image\/png;base64,)[^"]+("(?:[^{}]*)"sizes":"192x192")/, png192);
// manifest icons array: BOTH "any" and "maskable" 512x512 entries reuse the same 512 PNG
replaceAll(/("src":"data:image\/png;base64,)[^"]+("(?:[^{}]*)"sizes":"512x512")/g, png512);

fs.writeFileSync('index.html', html);
console.log(`Icons updated: ${replacements} references (192px ${png192.length}b64 chars, 512px ${png512.length}b64 chars)`);
