// 預覽操作手冊會被切成哪些條目、以及某個問句會挑到哪幾條（送進 prompt 的內容）
// 用法：node test/preview_manual.js [手册.md|手册.txt] ["問句"]
const fs = require('fs'), vm = require('vm'), path = require('path');
const file = process.argv[2] || path.join(require('os').tmpdir(), 'sample_manual.md');
const question = process.argv[3] || '如何點燈';
const raw = fs.readFileSync(file, 'utf8');
const props = {};
const ctx = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => props[k] = v, deleteProperty: k => delete props[k] }) },
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getId: () => 'F' }), openById: () => { throw new Error('no active'); } },
  Session: { getScriptTimeZone: () => 'Asia/Taipei' },
  Utilities: { formatDate: () => '', sleep() {} },
  CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
  UrlFetchApp: {}, HtmlService: { createHtmlOutput: x => x }, Logger: { log: console.log }, console
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'light.gs'), 'utf8'), ctx);
const R = e => vm.runInContext(e, ctx);

const name = path.basename(file).replace(/\.(md|txt|markdown)$/i, '');
const items = JSON.parse(R(`JSON.stringify(parseManualText(${JSON.stringify(raw)}, ${JSON.stringify(name)}))`));

console.log(`來源：${file}　原始 ${raw.length} 字 → 切成 ${items.length} 條\n`);
items.forEach((it, i) => console.log(`${String(i + 1).padStart(2)}. 【${it.topic}】\n` +
  it.body.split('\n').map(l => '     ' + l).join('\n') + '\n'));

R(`manualItems = function () { return ${JSON.stringify(items)}; }`);   // 預覽直接餵這份檔，不走來源偵測
const picked = R(`pickManual(${JSON.stringify(question)})`);
console.log(`=== 問句「${question}」實際進 prompt 的手冊部分（${picked.length} 字）===`);
console.log(picked);
