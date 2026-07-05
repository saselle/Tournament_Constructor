// Прокладка для `import * as XLSX from 'xlsx'`.
// xlsx-js-style подключена отдельным <script> в index.html и живёт в window.XLSX —
// эта прокладка просто переэкспортирует нужные части глобала как ES-модуль,
// вместо того чтобы esbuild пытался сделать настоящий require("xlsx") в браузере.
// Доступ к window.XLSX идёт лениво (в момент вызова, не в момент загрузки модуля),
// чтобы не зависеть от порядка <script>-тегов в index.html.
const getXlsx = () => (typeof window !== 'undefined' && window.XLSX) ? window.XLSX : {};
export const utils = new Proxy({}, { get: (_, prop) => getXlsx().utils[prop] });
export const write = (...args) => getXlsx().write(...args);
export default new Proxy({}, { get: (_, prop) => getXlsx()[prop] });
