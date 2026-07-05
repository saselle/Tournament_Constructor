// Вставляет свежесобранный dist/app_bundle.js в index.html между постоянными
// маркерами BEGIN_APP_BUNDLE / END_APP_BUNDLE. Маркеры не зависят от того, как
// минификатор называет переменные при очередной сборке — поэтому вставка
// стабильна между сборками (в отличие от поиска по содержимому бандла).
import fs from 'node:fs';

const html = fs.readFileSync('index.html', 'utf8');
const bundle = fs.readFileSync('dist/app_bundle.js', 'utf8');

const BEGIN = '<!-- BEGIN_APP_BUNDLE -->';
const END = '<!-- END_APP_BUNDLE -->';

const beginIdx = html.indexOf(BEGIN);
const endIdx = html.indexOf(END);
if (beginIdx === -1 || endIdx === -1) {
  throw new Error('BEGIN_APP_BUNDLE / END_APP_BUNDLE markers not found in index.html');
}

const before = html.slice(0, beginIdx + BEGIN.length);
const after = html.slice(endIdx);

const updated = `${before}<script>${bundle}</script>${after}`;
fs.writeFileSync('index.html', updated);
console.log(`index.html updated with fresh app bundle (${bundle.length} bytes)`);
