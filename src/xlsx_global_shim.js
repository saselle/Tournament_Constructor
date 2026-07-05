// Прокладка для `import * as XLSX from 'xlsx'`.
// xlsx-js-style подключена отдельным <script> в index.html и живёт в window.XLSX —
// эта прокладка просто переэкспортирует нужные части глобала как ES-модуль,
// вместо того чтобы esbuild пытался сделать настоящий require("xlsx") в браузере.
const xlsx = (typeof window !== 'undefined' && window.XLSX) ? window.XLSX : {};
export const utils = xlsx.utils;
export const write = xlsx.write;
export default xlsx;
