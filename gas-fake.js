// สภาพแวดล้อม Apps Script จำลอง (สำหรับทดสอบ Code.gs ในเครื่อง) + เว็บเซิร์ฟเวอร์ทดสอบ
const vm = require('vm'), fs = require('fs'), http = require('http'), crypto = require('crypto'), path = require('path');
const REPO = path.join(__dirname, '..'), PORT = Number(process.env.PORT || 8799);

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.d[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } o.push(row); } return o; }
  setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => this.sh.set(this.r + i, this.c + j, v))); return this; }
  setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) { const k = (this.r + i) + ',' + (this.c + j); f === '@' ? this.sh.txt.add(k) : this.sh.txt.delete(k); } return this; }
}
class Sheet {
  constructor(n) { this.n = n; this.d = []; this.txt = new Set(); }
  set(r, c, v) {
    while (this.d.length < r) this.d.push([]);
    const row = this.d[r - 1]; while (row.length < c) row.push('');
    if (this.txt.has(r + ',' + c)) v = (v instanceof Date) ? v : String(v);
    else if (typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v)) v = Number(v);   // จำลองการแปลงอัตโนมัติของ Sheets
    row[c - 1] = v;
  }
  _trim() { let last = this.d.length; while (last > 0 && this.d[last - 1].every(x => x === '' || x === undefined)) last--; return last; }
  getLastRow() { return this._trim(); }
  getMaxRows() { return Math.max(1000, this.d.length); }
  getDataRange() { const nr = Math.max(1, this._trim()), nc = Math.max(1, ...this.d.slice(0, nr).map(r => r.length)); return new Range(this, 1, 1, nr, nc); }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  deleteRow(i) { this.d.splice(i - 1, 1); const t = new Set(); /* เลื่อนรูปแบบข้อความตามแถว */ this.txt.forEach(k => { const [r, c] = k.split(',').map(Number); if (r < i) t.add(k); else if (r > i) t.add((r - 1) + ',' + c); }); this.txt = t; }
  setFrozenRows() {} getName() { return this.n; }
}
const sheets = {}, props = {}, cache = {}, drive = { files: {}, folders: {} };
let uid = 0;
class Blob { constructor(b, m, n) { this.b = b; this.m = m; this.n = n; } getBytes() { return this.b; } }
class File { constructor(blob, parent) { this.id = 'F' + String(++uid).padStart(26, 'x'); this.blob = blob; this.trashed = false; this.sharing = null; drive.files[this.id] = this; }
  getId() { return this.id; } getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view'; } setSharing(a, p) { this.sharing = [a, p]; return this; } setTrashed(t) { this.trashed = t; return this; } }
class Folder { constructor(n) { this.id = 'D' + String(++uid).padStart(26, 'x'); this.n = n; this.kids = []; this.files = []; drive.folders[this.id] = this; }
  getId() { return this.id; } getUrl() { return 'https://drive.google.com/drive/folders/' + this.id; }
  getFoldersByName(n) { const m = this.kids.filter(k => k.n === n); let i = 0; return { hasNext: () => i < m.length, next: () => m[i++] }; }
  createFolder(n) { const f = new Folder(n); this.kids.push(f); return f; }
  createFile(blob) { const f = new File(blob); this.files.push(f); return f; } }
const g = {
  console, Logger: { log: (...a) => console.log('[LOG]', ...a) }, Date, Math, JSON, Number, String, Object, Array, Set, Error,
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), getSheets: () => Object.values(sheets), deleteSheet: s => delete sheets[s.n], getName: () => 'ชีตทดสอบ', getUrl: () => 'https://docs.google.com/spreadsheets/d/TEST' }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, setProperties: o => Object.keys(o).forEach(k => { props[k] = String(o[k]); }), deleteProperty: k => { delete props[k]; } }) },
  CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Session: { getScriptTimeZone: () => 'Asia/Bangkok' },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) },
  DriveApp: { Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' }, Permission: { VIEW: 'VIEW' }, createFolder: n => new Folder(n), getFolderById: id => { if (!drive.folders[id]) throw new Error('not found'); return drive.folders[id]; }, getFileById: id => drive.files[id] },
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' },
    computeDigest: (a, s) => Array.from(crypto.createHash('sha256').update(String(s)).digest()).map(x => (x > 127 ? x - 256 : x)),
    base64Encode: x => Buffer.from(typeof x === 'string' ? x : x.map(v => (v + 256) % 256)).toString('base64'),
    base64Decode: s => Array.from(Buffer.from(s, 'base64')).map(x => (x > 127 ? x - 256 : x)),
    newBlob: (b, m, n) => new Blob(b, m, n), getUuid: () => crypto.randomUUID(),
    formatDate: (d, tz, f) => { const t = new Date(d.getTime() + 7 * 3600e3), p = (n, l = 2) => String(n).padStart(l, '0'); const M = { yyyy: t.getUTCFullYear(), yy: p(t.getUTCFullYear() % 100), MM: p(t.getUTCMonth() + 1), dd: p(t.getUTCDate()), HH: p(t.getUTCHours()), mm: p(t.getUTCMinutes()) };
      return f.replace(/'([^']*)'|yyyy|yy|MM|dd|HH|mm/g, (m, q) => (q !== undefined ? q : M[m])); }
  }
};
vm.createContext(g);
// const/class ระดับบนสุดใน vm แยก scope ต่อสคริปต์ จึงต่อท้ายเพื่อ export สิ่งที่ต้องใช้
vm.runInContext(fs.readFileSync(REPO + '/apps-script/Code.gs', 'utf8') + '\n;globalThis.__api={doPost,doGet,setup,seedDemo,UErr};', g, { filename: 'Code.gs' });
const A = g.__api;
const callGas = body => JSON.parse(A.doPost({ postData: { contents: body } }).s);
module.exports = { g, A, sheets, props, cache, drive, callGas };

if (require.main === module) {
  A.setup(); A.seedDemo();
  const cfgJs = `window.APP_CONFIG={API_URL:'http://localhost:${PORT}/exec'};`;
  http.createServer((q, r) => {
    const cors = { 'Access-Control-Allow-Origin': '*' };
    if (q.url.startsWith('/exec')) {
      if (q.method === 'GET') { r.writeHead(200, { ...cors, 'Content-Type': 'application/json' }); return r.end(A.doGet().s); }
      let b = ''; q.on('data', c => (b += c)); q.on('end', () => { r.writeHead(200, { ...cors, 'Content-Type': 'application/json' }); r.end(JSON.stringify(callGas(b))); });
      return;
    }
    const f = q.url === '/config.js' ? null : path.join(REPO, q.url === '/' ? 'index.html' : q.url.split('?')[0]);
    if (!f) { r.writeHead(200, { 'Content-Type': 'text/javascript' }); return r.end(cfgJs); }
    fs.readFile(f, (e, d) => { if (e) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': f.endsWith('.html') ? 'text/html; charset=utf-8' : 'text/plain' }); r.end(d); });
  }).listen(PORT, () => console.log('ready ' + PORT));
}
