const fs = require('fs'), vm = require('vm'), assert = require('assert');
const SRC = require('path').join(__dirname, '..', 'light.gs');

const HEADER = ['業務', '廟宇', '規格', '總燈數', '送燈時間', '軟體', '電腦', '備註'];
// 與真實分頁一致：A~H 共 8 欄，沒有正規日期輔助欄（國曆日期以〔國…〕標記併入備註）
const SEED = [
  HEADER,
  ['聖文', '石岡子乾元宮', '5*7 OLED琥珀色', 2112, '已送燈', '', '研華', '去年已換 5*7 全彩'],
  ['聖文', '桃園廣盛壇', '5*7 OLED琥珀色', 456, '國10/17前', '其他', '廟', ''],
  ['聖文', '桃園廣盛壇', '7*9 OLED琥珀色', 576, '', '其他', '廟', ''],
  ['甫穎', '仁武保安宮', '5*7 OLED琥珀色', 2475, '國10/22', '甫穎', '廟', ''],
  ['聖文', '田洋城隍廟', '5*7 OLED琥珀色', 1600, '國11/07前', '廟幫手', '廟', ''],
  ['聖文', '新化武廟', '4*5 OLED琥珀色', 5238, '國11/15前', '冠緯', '研華', ''],
  ['合計燈數', '', '', 12457, '', '', '', ''],
  ['士豪', '明倫三聖宮', '4*5 OLED', 7920, '12/10or12/17', '廟幫手', '研華', ''],
  ['士豪', '桃園慈護宮', '4*5 OLED', 55484, '國12/31前', '冠緯', '研華*3', ''],
  ['士豪', '三重聖佑宮', '5*7 OLED', 3728, '國115/03/15前', '廟管家', '研華', ''],
  ['士豪', '新竹城隍廟', '4*5 OLED', 15660, '國115/02/09前', '冠緯', '研華', ''],
  ['士豪', '金六結福德廟', '4*5 OLED', 4472, '國115/02/07前', '廟幫手', '廟', ''],
  ['士豪', '金六結福德廟', '5*7 OLED', 1407, '', '廟幫手', '廟', ''],
  ['士豪', '南崁五福宮', '4*5 OLED', 15600, '國115/02/13前', '家森', '研華', ''],
  ['趴一', '天成宮(北投)', '4*5 OLED', 108, '國12/31前', '冠宇', '廟', ''],
  ['趴一', '天成宮(中和)', '4*5 OLED', 108, '國12/31前', '冠宇', '廟', ''],
  ['合計燈數', '', '', 104487, '', '', '', '']
];

function makeSheet(rows, name) {
  const s = {
    _name: name || '', _rows: rows.map(r => r.slice()), _wrote: 0, _frozen: 0,
    _formulas: {},
    getName() { return s._name; },
    getDataRange() { return { getValues: () => s._rows }; },
    getMaxColumns() { return s._rows.reduce((a, r) => Math.max(a, r.length), 0); },
    getMaxRows() { return Math.max(s._rows.length, 1); },
    getLastRow() { return s._rows.length; },
    appendRow(v) { s._rows.push(v.slice()); s._wrote++; },
    insertColumnsAfter() {}, insertRowsAfter() {},
    setFrozenRows(n) { s._frozen = n; },
    protect() { const s2 = { setDescription: () => s2, setWarningOnly: () => s2 }; s._protected = true; return s2; },
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        setFontWeight() { return this; },
        clearContent() { s._cleared = (s._cleared || 0) + 1; return this; },
        setFormula(f) { s._formulas[r + ',' + c] = f; return this; },
        setValue(v) {
          while (s._rows.length < r) s._rows.push([]);
          const row = s._rows[r - 1];
          while (row.length < c) row.push('');
          row[c - 1] = v; s._wrote++;
        },
        setValues(vs) { vs.forEach((line, i) => line.forEach((v, j) => s.getRange(r + i, c + j).setValue(v))); }
      };
    }
  };
  return s;
}

function makeSS(masterName, masterRows) {
  const tabs = {};
  const master = makeSheet(masterRows, masterName);
  tabs[masterName] = master;
  const ss = {
    tabs, master,
    created: [],
    getSheetByName: n => tabs[n] || null,
    insertSheet: n => { const t = makeSheet([HEADER.slice()], n); tabs[n] = t; ss.created.push(n); return t; }
  };
  return ss;
}

let ssStub, sheet;
let cacheStore = {}, sent = [], ollamaCalls = [], geminiCalls = [];
function load() {
  ssStub = makeSS('光明燈管理', SEED);
  sheet = ssStub.master;
  const props = {};
  const ctx = {
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => props[k] = v, deleteProperty: k => delete props[k] }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getId: () => 'FAKE' }),
      openById: () => ssStub
    },
    Session: { getScriptTimeZone: () => 'Asia/Taipei' },
    __props: props,
    Utilities: { formatDate: (d, tz, f) => require('util').inspect(d).slice(0, 0) + `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`, sleep() {} },
    CacheService: { getScriptCache: () => ({
      get: k => (cacheStore[k] === undefined ? null : cacheStore[k]),
      put: (k, v) => { cacheStore[k] = String(v); },
      remove: k => { delete cacheStore[k]; }
    }) },
    UrlFetchApp: { fetch: (url, opt) => {
      const u = String(url);
      const ok = body => ({ getResponseCode: () => (ctx.HTTP_CODE || 200), getContentText: () => (typeof body === 'string' ? body : JSON.stringify(body)) });
      if (u.indexOf('/v2/bot/profile/') !== -1) return ok({ displayName: ctx.PROFILE_NAME || '' });
      if (u.indexOf('/v2/bot/message/reply') !== -1) {
        JSON.parse(opt.payload).messages.forEach(m => sent.push(m.text));
        return ok({});
      }
      if (u.indexOf('/api/chat') !== -1) {
        ollamaCalls.push(JSON.parse(opt.payload));
        return ok({ message: { content: ctx.OLLAMA_REPLY || 'OK 從 Ollama 來的回覆' } });
      }
      if (u.indexOf('generativelanguage') !== -1) {
        geminiCalls.push(u);
        return ok({ candidates: [{ content: { parts: [{ text: ctx.GEMINI_TEXT || JSON.stringify({ action: 'READ', details: {} }) }] } }] });
      }
      return ok('{}');
    } },
    HtmlService: { createHtmlOutput: x => ({ content: x }) },
    DocumentApp: {
      openById: id => {
        const f = (ctx.FOLDER || []).find(x => x.id === id);
        if (!f) throw new Error('no doc ' + id);
        return { getBody: () => ({ getText: () => f.text }) };
      }
    },
    DriveApp: {
      getFolderById: () => {
        const files = (ctx.FOLDER || []).filter(f => !f.standalone);
        let i = 0;
        return { getFiles: () => ({ hasNext: () => i < files.length, next: () => {
          const f = files[i++];
          return {
            getName: () => f.name, getMimeType: () => f.mime, getId: () => f.id,
            getBlob: () => ({ getDataAsString: () => f.text })
          };
        } }) };
      }
    },
    Logger: { log() {} }, console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(SRC, 'utf8'), ctx, { filename: 'light.gs' });
  return ctx;
}

const ctx = load();
const run = expr => vm.runInContext(expr, ctx);
let pass = 0, fail = 0;
function reset() {
  Object.keys(ssStub.tabs).forEach(k => delete ssStub.tabs[k]);   // 同一個物件，不能整包換掉（閉包會抓到舊的）
  ssStub.tabs['光明燈管理'] = sheet; ssStub.created = [];
  sheet._rows = SEED.map(r => r.slice()); sheet._wrote = 0;
  sheet._formulas = {}; sheet._cleared = 0; sheet._protected = false;
  run('CONFIG.DEFAULT_AGENT = "聖文"; CONFIG.AGENT_FROM_LINE_PROFILE = false; CONFIG.SPREADSHEET_ID = "FAKE";');
  cacheStore = {}; sent = []; ollamaCalls = []; geminiCalls = [];
  ctx.OLLAMA_REPLY = ''; ctx.GEMINI_TEXT = ''; ctx.HTTP_CODE = 200; delete ctx.__props.OLLAMA_API_KEY;
  ctx.FOLDER = [];
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = ""; CONFIG.CHAT.MANUAL_DOC_ID = ""; CONFIG.CHAT.MANUAL_MAX_FILES = 10;');
  ctx.__props.REPLY_MODE = ''; delete ctx.__props.USER_IDENTS;
  run('CONFIG.REPLY_MODE = ""; CONFIG.OLLAMA_API_KEY = "YOUR_OLLAMA_API_KEY";');
  run('CONFIG.LLM.PROVIDER = "GEMINI"; CONFIG.LLM.CHAT_PROVIDER = ""; CONFIG.LLM.OLLAMA_MODEL = "gpt-oss:20b";');
  run('CONFIG.LLM.OLLAMA_BASE_URL = "https://ollama.com"; CONFIG.CHAT.MAX_REPLY_CHARS = 600;');
  run('CONFIG.CHAT.MANUAL_DOC_ID = ""; CONFIG.CHAT.MANUAL_SHEET = "操作手冊";'); run('CONFIG.AGENT_FROM_LINE_PROFILE = false');
  ctx.__NEXT_AI = null; ctx.PROFILE_NAME = '';
}
function t(name, fn) {
  reset();
  try { fn(); pass++; console.log('  ✓ ' + name); }
  catch (e) { fail++; console.log('  ✗ ' + name + '\n      ' + e.message); }
}

console.log('\n【日期正規化 parseDateKey】');
[['國10/17前', '2026-10-17'], ['國115-02/09前', '2026-02-09'], ['國115/02/09前', '2026-02-09'],
 ['12/10or12/17', '2026-12-10'], ['10月15日', '2026-10-15'], ['已送燈', ''],
 ['國115/03/15前', '2026-03-15'], ['2026-11-07', '2026-11-07'], ['國12/31前', '2026-12-31'],
 ['', ''], [null, ''], ['待定', '']].forEach(([raw, want]) => {
  t(`"${raw}" -> ${want || '(空)'}`, () => assert.strictEqual(run(`parseDateKey(${JSON.stringify(raw)})`), want));
});

console.log('\n【小計列與比對】');
t('isSummaryRow 抓到合計燈數', () => assert.ok(run('isSummaryRow(["合計燈數","","",12457,"","","",""])')));
t('isSummaryRow 不误判一般列', () => assert.ok(!run('isSummaryRow(["聖文","石岡子乾元宮","5*7 OLED琥珀色",2112,"已送燈","","研華",""])')));
t('規格 5*7 / 5×7 / 5 7 相同', () => assert.ok(run('matchSpec("5*7 OLED琥珀色","5×7 oled 琥珀色")')));
t('規格 4*5 不等於 5*7', () => assert.ok(!run('matchSpec("4*5 OLED","5*7 OLED")')));
t('廟名 天成宮 可中找到天成宮(北投)', () => assert.ok(run('matchTemple("天成宮(北投)","天成宮")')));

console.log('\n【新增 CREATE】');
t('重複登記被擋下', () => {
  const before = sheet._rows.length;
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '聖文', temple: '石岡子乾元宮', spec: '5*7 OLED琥珀色', total_count: 2112, delivery_text: '已送燈', software: null, computer: '研華' } })}, "聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞")`);
  assert.ok(/疑似重複登記/.test(r), r);
  assert.ok(/第 2 列/.test(r), '應指出既有列號：' + r);
  assert.strictEqual(sheet._rows.length, before, '不應寫入');
});
t('強制新增可繞過', () => {
  const before = sheet._rows.length;
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '聖文', temple: '石岡子乾元宮', spec: '5*7 OLED琥珀色', total_count: 300, delivery_text: '國11/20前', software: '廟幫手', computer: '研華' } })}, "強制新增")`);
  assert.ok(/新增成功/.test(r), r);
  assert.strictEqual(sheet._rows.length, before + 1);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(JSON.stringify(last), JSON.stringify(['聖文', '石岡子乾元宮', '5*7 OLED琥珀色', 300, '國11/20前', '廟幫手', '研華', '〔國2026-11-20〕']));
});
t('一般新增寫入 8 欄', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '冠宇', temple: '行天宮', spec: '5*7 OLED琥珀色', total_count: 880, delivery_text: '國115/02/09前', software: '冠宇', computer: '研華' } })}, "x")`);
  assert.ok(/新增成功/.test(r), r);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[3], 880); assert.strictEqual(last[7], '〔國2026-02-09〕');
});
t('缺廟名與規格時不寫入', () => {
  const before = sheet._rows.length;
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '聖文', temple: null, spec: null, total_count: 100 } })}, "x")`);
  assert.ok(/新增失敗/.test(r), r);
  assert.strictEqual(sheet._rows.length, before);
});

console.log('\n【修改 UPDATE】');
t('同廟多規格 -> 請指明規格', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '金六結福德廟', spec: null, total_count: 5000, delivery_text: null, software: null, computer: null } })}, "x")`);
  assert.ok(/需要指明規格/.test(r), r);
  assert.ok(/4\*5 OLED/.test(r) && /5\*7 OLED/.test(r), '應列出兩個規格：' + r);
});
t('指明規格 -> 只改該列', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '金六結福德廟', spec: '5*7 OLED', total_count: 1500, delivery_text: '國115/03/01前', software: null, computer: '廟' } })}, "x")`);
  assert.ok(/修改成功/.test(r), r);
  assert.ok(/總燈數：1,407 盞 → 1,500 盞/.test(r), r);
  const row = sheet._rows[13];
  assert.strictEqual(row[0], '士豪', '不應動到業務欄：' + JSON.stringify(row));
  assert.strictEqual(row[3], 1500);
  assert.strictEqual(row[4], '國115/03/01前');
  assert.strictEqual(row[7], '〔國2026-03-01〕', '備註應同步國曆標記：' + row[7]);
});
t('數值相同 -> 不做更新', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '新化武廟', spec: '4*5 OLED琥珀色', total_count: 5238, delivery_text: null, software: null, computer: null } })}, "x")`);
  assert.ok(/未做更新/.test(r), r);
});
t('找不到廟宇 -> 提示既有規格', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '不存在宮', spec: '9*9 LED', total_count: 1 } })}, "x")`);
  assert.ok(/修改失敗/.test(r), r);
});
t('修改不會動到合計列', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '合計燈數', spec: null, total_count: 999 } })}, "x")`);
  assert.ok(/修改失敗/.test(r), r);
  assert.strictEqual(sheet._rows[7][3], 12457, '合計列被改到：' + JSON.stringify(sheet._rows[6]));
});

console.log('\n【查詢 READ】');
t('查 天成宮 -> 北投+中和 兩筆，合計 216', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: { temple: '天成宮' } })}, "x")`);
  assert.ok(/共 2 筆/.test(r), r);
  assert.ok(/合計 216 盞/.test(r), r);
});
t('查軟體 冠緯 -> 3 筆', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: { software: '冠緯' } })}, "x")`);
  assert.ok(/共 3 筆/.test(r), r);
});
t('查規格 5*7 -> 抓到所有 5*7', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: { spec: '5*7' } })}, "x")`);
  assert.ok(/共 6 筆/.test(r), r);
});
t('無條件 -> 顯示最新且不含合計', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: {} })}, "x")`);
  assert.ok(/最近 12 筆/.test(r), r);
  assert.ok(/共 15 筆/.test(r), r);
  assert.ok(!/合計燈數/.test(r), '不應列出合計列：' + r);
});
t('查無結果', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: { temple: '歐多貝聖宮' } })}, "x")`);
  assert.ok(/查詢無結果/.test(r), r);
});

console.log('\n【指令分流】');
t('/help 命中說明', () => assert.ok(run('isHelpCommand("/help")') && /使用方式/.test(run('handleHelpCommand()'))));
t('/model 不被 /help 攔截', () => assert.ok(!run('isHelpCommand("/model")')));
t('一般登記文字走 Gemini', () => assert.ok(!run('isHelpCommand("聖文-石岡子乾元宮 5*7 2112盞")')));

console.log('\n【右側公式欄共存】');
t('右側 I 欄有「合計燈數」標籤 -> 一般列不被誤判為小計', () => {
  assert.ok(!run('isSummaryRow(["士豪","南崁五福宮","4*5 OLED",15600,"國115/02/13前","家森","研華","","合計燈數",99999])'));
});
t('左側 A 欄仍是小計 -> 正確判斷', () => {
  assert.ok(run('isSummaryRow(["合計燈數","","",12457,"","","","","",12457])'));
});
t('每列右側都掛公式欄 -> 查詢筆數不變', () => {
  sheet._rows = SEED.map(r => r.concat(['合計燈數', 5000]));
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: { temple: '天成宮' } })}, "x")`);
  assert.ok(/共 2 筆/.test(r), r);
  assert.ok(/合計 216 盞/.test(r), r);
});
t('右側有公式欄 -> 新增只寫 A~H 八欄', () => {
  sheet._rows = SEED.map(r => r.concat(['小計', 1]));
  run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '冠宇', temple: '寶林宮', spec: '5*7 OLED', total_count: 66, delivery_text: '國11/30前', software: null, computer: null } })}, "x")`);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last.length, 8, '應只有 A~H 八欄：' + JSON.stringify(last));
  assert.strictEqual(last[7], '〔國2026-11-30〕');
});
t('右側有公式欄 -> 修改仍只動 A~H', () => {
  sheet._rows = SEED.map(r => r.concat(['小計', 12457]));
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: { agent: null, temple: '田洋城隍廟', spec: '5*7 OLED琥珀色', total_count: 1700, delivery_text: null, software: null, computer: null } })}, "x")`);
  assert.ok(/修改成功/.test(r), r);
  const row = sheet._rows[5];
  assert.strictEqual(row[3], 1700);
  assert.strictEqual(row[8], '小計', '右側公式欄標籤不應被覆蓋');
  assert.strictEqual(row[9], 12457, '右側公式值不應被改動');
});

console.log('\n【只改單一欄位，其他資料不遺失】');
const detail = o => Object.assign({ agent: null, temple: null, spec: null, total_count: null, delivery_text: null, software: null, computer: null }, o);
t('修改 新化武廟 總燈數變成5238盞（未提規格，該廟僅一列）', () => {
  const before = sheet._rows[6].slice();
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: detail({ temple: '新化武廟', total_count: 5300 }) })}, "修改新化武廟 總燈數變成5300盞")`);
  assert.ok(/修改成功/.test(r), r);
  assert.ok(/第 7 列/.test(r), r);
  assert.ok(/總燈數：5,238 盞 → 5,300 盞/.test(r), r);
  const after = sheet._rows[6];
  assert.strictEqual(after[3], 5300);
  [0, 1, 2, 4, 5, 6, 7].forEach(i => assert.strictEqual(after[i], before[i], `第 ${i + 1} 欄被動到：` + JSON.stringify([before, after])));
});
t('只改軟體 -> 燈數與電腦不變', () => {
  const before = sheet._rows[10].slice();
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: detail({ temple: '三重聖佑宮', spec: '5*7', software: '冠緯' }) })}, "x")`);
  assert.ok(/軟體：廟管家 → 冠緯/.test(r), r);
  const after = sheet._rows[10];
  assert.strictEqual(after[3], before[3], '燈數不應變');
  assert.strictEqual(after[6], before[6], '電腦不應變');
  assert.strictEqual(after[4], before[4], '送燈日期不應變');
});
t('未給規格且該廟多規格 -> 一欄都不寫', () => {
  const before = JSON.stringify(sheet._rows);
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: detail({ temple: '桃園廣盛壇', total_count: 999 }) })}, "x")`);
  assert.ok(/需要指明規格/.test(r), r);
  assert.strictEqual(JSON.stringify(sheet._rows), before, '不應有任何寫入');
});
t('規格改成 -> 只換規格欄', () => {
  const before = sheet._rows[6].slice();
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: detail({ temple: '新化武廟', spec: '5*7 OLED琥珀色' }) })}, "新化武廟 規格改成 5*7 OLED琥珀色")`);
  assert.ok(/修改成功/.test(r), r);
  assert.ok(/規格：4\*5 OLED琥珀色 → 5\*7 OLED琥珀色/.test(r), r);
  const after = sheet._rows[6];
  assert.strictEqual(after[3], before[3], '燈數不應變');
  assert.strictEqual(after[4], before[4], '日期不應變');
});
t('同廟同規格兩列 + 有給規格 -> 提示再用業務區分，不亂寫', () => {
  run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: detail({ agent: '士豪', temple: '金六結福德廟', spec: '4*5 OLED', total_count: 100, delivery_text: '國115/04/01前', software: '廟幫手', computer: '廟' }) })}, "強制新增")`);
  const before = JSON.stringify(sheet._rows);
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: detail({ temple: '金六結福德廟', spec: '4*5 OLED', total_count: 777 }) })}, "x")`);
  assert.ok(/同一規格也有多筆/.test(r), r);
  assert.strictEqual(JSON.stringify(sheet._rows).length, before.length, '不應寫入');
});

console.log('\n【回覆策略：只回「會動到資料」的訊息】');
const flow = t2 => run(`decideFlow(${JSON.stringify(t2)}).mode`);
const handlerOf = t2 => run(`(decideFlow(${JSON.stringify(t2)})||{}).handler`);
t('查詢語句 -> 沉默（連 AI 都不調）', () => assert.strictEqual(flow('查一下 天成宮 有哪些燈'), 'silent'));
t('閒聊「今天天氣如何」-> 沉默', () => assert.strictEqual(flow('今天天氣如何'), 'silent'));
t('問候「你是谁」-> 沉默', () => assert.strictEqual(flow('你是谁'), 'silent'));
t('空訊息 -> 沉默', () => assert.strictEqual(flow('   '), 'silent'));
t('帶喚醒字的完整登記句 -> 送 AI', () => assert.strictEqual(flow('小幫手 聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前 軟體其他 電腦研華'), 'ai'));
t('帶喚醒字的修改句 -> 送 AI', () => assert.strictEqual(flow('小幫手 修改 新化武廟 總燈數變成5238盞'), 'ai'));
t('未帶喚醒字的登記句 -> 沉默（不花配額）', () => assert.strictEqual(flow('聖文-石岡子乾元宮 5*7 2112盞'), 'silent'));
t('喚醒字在中間也能辨識', () => assert.strictEqual(run('stripWakeWord("請小幫手協助登記 聖文-田洋城隍廟 5*7 1600盞")').indexOf('請'), 0));
const idt = t2 => run(`identityLabel(parseIdentityCommand(${JSON.stringify(t2)}))`);
t('身分綁定：各種寫法與顯示標籤', () => {
  assert.strictEqual(idt('我是聖文'), '聖文');
  assert.strictEqual(idt('我公司 亞盛燈業有限公司'), '亞盛燈業有限公司');
  assert.strictEqual(idt('業務窗口 冠緯'), '冠緯');
  assert.strictEqual(idt('我叫 李小華'), '李小華');
  assert.strictEqual(idt('我公司 亞盛燈業 我叫 李小華'), '亞盛燈業-李小華');
  assert.strictEqual(idt('我是 亞盛燈業 的 李小華'), '亞盛燈業-李小華');
});
t('看起來像登記的內容絕不當綁定', () => {
  [
    '我是 石岡子乾元宮 5*7 OLED 2112盞',
    '我是聖文 石岡子乾元宮',
    '聖文-石岡子乾元宮 5*7 2112盞 國10/17前',
    '今天天氣如何',
    '我公司 亞盛燈業 新增 5*7 OLED 2112盞'
  ].forEach(txt => assert.strictEqual(run(`parseIdentityCommand(${JSON.stringify(txt)})`), null, '應回傳 null：' + txt));
});
t('/help 仍會回覆', () => { assert.strictEqual(flow('/help'), 'command'); assert.strictEqual(handlerOf('/help'), 'help'); });
t('/model 仍會回覆', () => { assert.strictEqual(flow('/model'), 'command'); assert.strictEqual(handlerOf('/model'), 'model'); });
t('意圖把關：CREATE/UPDATE 才回', () => {
  assert.strictEqual(run('shouldReplyFor("CREATE")'), true);
  assert.strictEqual(run('shouldReplyFor("update")'), true);
  assert.strictEqual(run('shouldReplyFor("READ")'), false);
  assert.strictEqual(run('shouldReplyFor("OTHER")'), false);
  assert.strictEqual(run('shouldReplyFor("")'), false);
});
t('REPLY_MODE=ALL 可全部回覆', () => {
  run('CONFIG.REPLY_MODE = "ALL"');
  assert.strictEqual(flow('查一下 天成宮 有哪些燈'), 'ai');
  assert.strictEqual(run('shouldReplyFor("READ")'), true);
});
t('REPLY_MODE=off 不影響（非 ALL 值）', () => {
  run('CONFIG.REPLY_MODE = "WRITE"');
  assert.strictEqual(flow('查一下 天成宮 有哪些燈'), 'silent');
  assert.strictEqual(run('shouldReplyFor("READ")'), false);
});

console.log('\n【喚醒字、必填攔截與補件（端到端 doPost）】');
vm.runInContext(`
  var AI_INPUTS = [];
  var _origAI = analyzeMessageWithGemini;
  analyzeMessageWithGemini = function(t) { AI_INPUTS.push(t); return (__NEXT_AI === null || __NEXT_AI === undefined) ? _origAI(t) : __NEXT_AI; };
`, ctx);
const D = o => Object.assign({ agent: null, temple: null, spec: null, total_count: null, delivery_text: null, software: null, computer: null }, o);
const post = (text, uid) => { vm.runInContext('AI_INPUTS.length = 0', ctx); run('doPost')({ postData: { contents: JSON.stringify({ events: [{ type: 'message', message: { type: 'text', text }, replyToken: 'RT', source: { type: 'user', userId: uid || 'U_TEST' } }] }) } }); };
const ai = o => { ctx.__NEXT_AI = o; };

t('未帶喚醒字 -> 完全不回、不調 AI', () => {
  ai({ action: 'CREATE', details: D({ temple: 'X宮' }) });
  post('聖文 石岡子乾元宮 5*7 OLED琥珀色 2112盞');
  assert.strictEqual(sent.length, 0, '不應回訊');
  assert.strictEqual(vm.runInContext('AI_INPUTS.length', ctx), 0, '不應花 Gemini 配額');
});
t('缺欄追問後 10 分鐘內補件 -> 合併草稿並寫入（兩步驟）', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '媽祖廟', spec: '財神燈' }) });
  post('小幫手 我要登記，媽祖廟要新增財神燈');
  const rowsAfterAsk = sheet._rows.length;
  assert.strictEqual(sent.length, 1, '應回覆補件提示');
  assert.ok(/資料還沒齊全，暫未登記/.test(sent[0]), sent[0]);
  assert.ok(/還需要：總燈數/.test(sent[0]), sent[0]);
  assert.ok(/可以直接補：小幫手 2112盞/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows.length, rowsAfterAsk, '不應寫入試算表');
  assert.ok(run('loadDraft("U_TEST")'), '草稿應已暫存');
  sent.length = 0;
  ai({ action: 'CREATE', details: D({ temple: '媽祖廟', spec: '財神燈', total_count: 500 }) });
  post('小幫手 500盞');
  const joined = vm.runInContext('AI_INPUTS[0]', ctx);
  assert.ok(/補充：500盞/.test(joined), '應把前一句串起來給 AI：' + joined);
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[0], '聖文'); assert.strictEqual(last[1], '媽祖廟'); assert.strictEqual(last[3], 500);
  assert.strictEqual(run('loadDraft("U_TEST")'), null, '寫入後草稿應清掉');
});
t('（同一測試的前段重複檢查）缺總燈數 -> 不寫表', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '媽祖廟', spec: '財神燈' }) });
  post('小幫手 我要登記，媽祖廟要新增財神燈');
  const before = sheet._rows.length;
  assert.strictEqual(sent.length, 1, '應回覆補件提示');
  assert.ok(/資料還沒齊全，暫未登記/.test(sent[0]), sent[0]);
  assert.ok(/還需要：總燈數/.test(sent[0]), sent[0]);
  assert.ok(/可以直接補：小幫手 2112盞/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows.length, before, '不應寫入試算表');
});

t('沒寫業務 -> 上預設承攬商「聖文」，不會卡也問不到暱稱', () => {
  ctx.PROFILE_NAME = '小明🔥';
  ai({ action: 'CREATE', details: D({ temple: '武勝宮', spec: '5*7 OLED', total_count: 200 }) });
  post('小幫手 登記 武勝宮 5*7 OLED 200盞');
  assert.ok(/業務預設帶入：聖文/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '聖文');
});
t('綁定優先於預設（聖文員工不會被寫成別家）', () => {
  post('小幫手 我公司 名典 我叫 李小華');
  sent.length = 0;
  ai({ action: 'CREATE', details: D({ temple: '武勝宮', spec: '5*7 OLED', total_count: 200 }) });
  post('小幫手 登記 武勝宮 5*7 OLED 200盞');
  assert.ok(/業務自動帶入（您的綁定）：名典-李小華/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '名典-李小華');
});
t('DEFAULT_AGENT 設空 -> 沒綁定也沒寫業務時才回問', () => {
  run('CONFIG.DEFAULT_AGENT = ""');
  ctx.PROFILE_NAME = '小明🔥';
  ai({ action: 'CREATE', details: D({ temple: '武勝宮', spec: '5*7 OLED', total_count: 200 }) });
  post('小幫手 登記 武勝宮 5*7 OLED 200盞', 'U_NO_BIND_2');
  assert.ok(/還需要：業務／公司名/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows.length, 18, '不應寫入任何一列');
});
t('開 AGENT_FROM_LINE_PROFILE=ON 且無預設 -> 才自動帶 LINE 暱稱', () => {
  run('CONFIG.AGENT_FROM_LINE_PROFILE = true; CONFIG.DEFAULT_AGENT = ""');
  ctx.PROFILE_NAME = '小明🔥';
  ai({ action: 'CREATE', details: D({ temple: '武勝宮', spec: '5*7 OLED', total_count: 200 }) });
  post('小幫手 登記 武勝宮 5*7 OLED 200盞');
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
  assert.ok(/自動帶入 LINE 暱稱「小明🔥」/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '小明🔥');
  run('CONFIG.AGENT_FROM_LINE_PROFILE = false');
});
t('「我是王大仁」綁定姓名，之後優先套用綁定名', () => {
  ctx.PROFILE_NAME = '小明🔥';
  post('小幫手 我是王大仁');
  assert.ok(/已記住您的業務身分：「王大仁」/.test(sent[0]), sent[0]);
  sent.length = 0;
  ai({ action: 'CREATE', details: D({ temple: '保安宮', spec: '4*5 OLED', total_count: 300 }) });
  post('小幫手 登記 保安宮 4*5 OLED 300盞');
  assert.ok(/業務自動帶入（您的綁定）：王大仁/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '王大仁');
});
t('「我是OO」開頭的整串登記不會被當成綁定', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '九曲宮', spec: '5*7 OLED', total_count: 2112 }) });
  post('小幫手 我是聖文 九曲宮 5*7 OLED 2112盞');
  assert.ok(!/已記住您的業務身分/.test(sent[0] || ''), sent[0]);
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
});
t('全部拿不到（無預設、無綁定、無暱稱）-> 請補業務／公司名', () => {
  run('CONFIG.DEFAULT_AGENT = ""');
  ctx.PROFILE_NAME = '';
  ai({ action: 'CREATE', details: D({ temple: '廣正宮', spec: '5*7 OLED', total_count: 100 }) });
  post('小幫手 登記 廣正宮 5*7 OLED 100盞', 'U_NO_NAME');
  assert.ok(/還需要：業務／公司名/.test(sent[0]), sent[0]);
  assert.ok(/我是 聖文/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows.length, 18, '不應寫入任何一列');
});
t('修改類不需喚醒字以外條件：帶了喚醒字就處理', () => {
  ai({ action: 'UPDATE', details: D({ temple: '新化武廟', total_count: 5300 }) });
  post('小幫手 修改 新化武廟 總燈數變成5300盞');
  assert.ok(/修改成功/.test(sent[0]), sent[0]);
});
t('查詢類即使帶喚醒字也不回訊', () => {
  ai({ action: 'READ', details: D({ temple: '天成宮' }) });
  post('小幫手 查一下 天成宮 有哪些燈');
  assert.strictEqual(sent.length, 0, '查詢不應回訊');
});
t('綁公司+人員 -> 業務欄顯示「公司-人員」', () => {
  post('小幫手 我公司 亞盛燈業 我叫 李小華');
  assert.ok(/「亞盛燈業-李小華」/.test(sent[0]), sent[0]);
  sent.length = 0;
  ai({ action: 'CREATE', details: D({ temple: '保安宮', spec: '4*5 OLED', total_count: 120 }) });
  post('小幫手 登記 保安宮 4*5 OLED 120盞');
  assert.ok(/業務自動帶入（您的綁定）：亞盛燈業-李小華/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '亞盛燈業-李小華');
});
t('綁定後改公司名 -> 直接蓋掉舊名', () => {
  post('小幫手 我是王大仁');
  post('小幫手 我公司 亞盛燈業');
  assert.ok(/亞盛燈業/.test(sent[sent.length - 1]), sent[sent.length - 1]);
  assert.strictEqual(run('JSON.parse(__props.USER_IDENTS).U_TEST.company'), '亞盛燈業');
});
t('REPLY_MODE=ALL 時不用喚醒字也處理', () => {
  run('CONFIG.REPLY_MODE = "ALL"');
  ai({ action: 'READ', details: D({ temple: '天成宮' }) });
  post('查一下 天成宮 有哪些燈');
  assert.ok(/查詢結果/.test(sent[0]), sent[0]);
});

console.log('\n【業務／公司名 批次改名】');
t('dryRun 只預覽不寫入', () => {
  const before = JSON.stringify(sheet._rows);
  const r = run(`renameAgent("聖文", "聖文企業", true)`);
  assert.strictEqual(sheet._rows.length, 18);
  assert.strictEqual(JSON.stringify(sheet._rows), before, 'dryRun 不應寫入');
  assert.ok(/共 5 列/.test(r), r);
  assert.ok(/第 2 列：聖文 → 聖文企業（石岡子乾元宮）/.test(r), r);
});
t('真的改名 -> 4 列更新，合計列不碰', () => {
  const r = run(`renameAgent("聖文", "聖文企業")`);
  assert.ok(/已完成/.test(r), r);
  [1, 2, 3, 5, 6].forEach(i => assert.strictEqual(sheet._rows[i][0], '聖文企業', '第 ' + (i + 1) + ' 列'));
  assert.strictEqual(sheet._rows[4][0], '甫穎', '其他人的列不應被改');
  assert.strictEqual(sheet._rows[7][0], '合計燈數', '合計列不應被改');
  assert.strictEqual(sheet._rows[7][3], 12457, '合計數值不應被改');
});
t('已綁定的身分名稱也一起換', () => {
  run(`bindIdentity("U_B", { company: "聖文" })`);
  run(`renameAgent("聖文", "聖文企業")`);
  assert.strictEqual(run('identityLabel(readIdentities()["U_B"])'), '聖文企業');
});
t('LINE 端：業務改成 OO（業務欄當目標而非定位）', () => {
  ai({ action: 'UPDATE', details: D({ agent: '王大川', temple: '石岡子乾元宮' }) });
  post('小幫手 修改 石岡子乾元宮 業務改成王大川');
  assert.ok(/業務：聖文 → 王大川/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[1][0], '王大川');
  assert.strictEqual(sheet._rows[1][3], 2112, '燈數不應被改');
});
t('一般修改會用業務欄共同過濾（業務不符就不亂改）', () => {
  ai({ action: 'UPDATE', details: D({ agent: '冠宇', temple: '天成宮', total_count: 5 }) });
  post('小幫手 修改 天成宮 冠宇 總燈數變成5盞');
  assert.ok(/修改失敗/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[15][3], 108, '不應寫入');
});
t('業務相符 + 指明分館 -> 只改那一列', () => {
  ai({ action: 'UPDATE', details: D({ agent: '趴一', temple: '天成宮(北投)', total_count: 118 }) });
  post('小幫手 修改 趴一 天成宮(北投) 總燈數變成118盞');
  assert.ok(/修改成功/.test(sent[0]), sent[0]);
  assert.strictEqual(sheet._rows[15][3], 118);
  assert.strictEqual(sheet._rows[16][3], 108, '中和館不應被改');
});

console.log('\n【備註欄組裝 buildRemark / stripDateTag】');
t('只有日期沒有說明 -> 標記獨存', () => assert.strictEqual(run(`buildRemark(null, '國10/17前')`), '〔國2026-10-17〕'));
t('有說明 -> 標記置頂、說明照原文', () => assert.strictEqual(run(`buildRemark('分兩批送', '國10/17前')`), '〔國2026-10-17〕 分兩批送'));
t('重複組裝不會堆疊標記', () => {
  const once = run(`buildRemark('分兩批送', '國10/17前')`);
  assert.strictEqual(run(`buildRemark(${JSON.stringify(once)}, '國10/17前')`), once);
});
t('換日期 -> 舊標記被替換', () => assert.strictEqual(run(`buildRemark('〔國2026-10-17〕 分兩批送', '國11/05前')`), '〔國2026-11-05〕 分兩批送'));
t('已送燈/農曆 -> 不插標記，只留說明', () => {
  assert.strictEqual(run(`buildRemark('改天再確認', '已送燈')`), '改天再確認');
  assert.strictEqual(run(`buildRemark('改天再確認', '農曆10/17前')`), '改天再確認');
});
t('DATE_TAG_IN_REMARK=OFF -> 完全不插標記', () => {
  run('CONFIG.REMARK_DATE_TAG = false');
  assert.strictEqual(run(`buildRemark('分兩批送', '國10/17前')`), '分兩批送');
  run('CONFIG.REMARK_DATE_TAG = true');
});
t('說明裡已有同一日期 -> 不重複插', () => assert.strictEqual(run(`buildRemark('國曆2026-10-17 送', '國10/17前')`), '國曆2026-10-17 送'));

console.log('\n【備註欄與曆別（國曆/民國/農曆）】');
t('明示備註才會寫入 H 欄（並自動前置國曆標記）', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '永樂宮', spec: '5*7 OLED', total_count: 88, delivery_text: '國11/25前', remark: '分兩批送，第二批國12/01前' }) });
  post('小幫手 登記 永樂宮 5*7 OLED 88盞 國11/25前 備註：分兩批送，第二批國12/01前');
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[7], '〔國2026-11-25〕 分兩批送，第二批國12/01前', '整列=' + JSON.stringify(last) + '｜回覆=' + sent[0]);
  assert.strictEqual(last.length, 8);
});
t('沒提備註也沒日期 -> H 欄留空，不拿整句原文亂填', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '樂善宮', spec: '5*7 OLED', total_count: 10 }) });
  post('小幫手 登記 樂善宮 5*7 OLED 10盞');
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][7], '');
});
t('修改可單獨改備註', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: D({ temple: '石岡子乾元宮', spec: '5*7', remark: '已改全彩' }) })}, "x")`);
  assert.ok(/備註：去年已換 5\*7 全彩 → 已改全彩/.test(r), r);
  assert.strictEqual(sheet._rows[1][7], '已改全彩');
  assert.strictEqual(sheet._rows[1][3], 2112, '燈數不應被改');
});
t('查詢會列出備註', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'READ', details: D({ temple: '石岡子乾元宮' }) })}, "x")`);
  assert.ok(/備註：/.test(r), r);
});
t('民國 115/116 -> 西元 2026/2027', () => {
  assert.strictEqual(run(`parseDateKey("國115/03/15前")`), '2026-03-15');
  assert.strictEqual(run(`parseDateKey("民國116/01/02")`), '2027-01-02');
  assert.strictEqual(run(`parseDateKey("西元2027/06/30")`), '2027-06-30');
});
t('農曆不換算（H 欄留空，E 欄仍保留原文）', () => {
  ['', '農10/17', '農曆十月十七', '舊曆11/05', '陰曆12/01前'].forEach(t2 => {
    assert.strictEqual(run(`parseDateKey(${JSON.stringify(t2)})`), '', '應留空：' + t2);
  });
});
t('纯國曆與只有月日也照舊換算', () => {
  assert.strictEqual(run(`parseDateKey("國曆10/17前")`), '2026-10-17');
  assert.strictEqual(run(`parseDateKey("12/25")`), '2026-12-25');
});
t('登記一句帶農曆 -> E 欄原樣、H 欄空，不會推出錯日期', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '聖安宮', spec: '5*7 OLED', total_count: 50, delivery_text: '農曆10/17前' }) });
  post('小幫手 登記 聖安宮 5*7 OLED 50盞 農曆10/17前');
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[4], '農曆10/17前', 'E 欄應保留原寫法：' + JSON.stringify(last));
  assert.strictEqual(last[7], '', '備註不應亂推日期：' + last[7]);
});

console.log('\n【分頁自動建立與提示】');
t('分頁不存在 -> 錯誤訊息導向 setupLightSheet()', () => {
  delete ssStub.tabs['光明燈管理'];
  ai({ action: 'CREATE', details: D({ temple: '武勝宮', spec: '5*7 OLED', total_count: 200 }) });
  post('小幫手 登記 武勝宮 5*7 OLED 200盞');
  assert.ok(/setupLightSheet/.test(sent[0]), sent[0]);
});
t('setupLightSheet() 自動建立主表 + 同步公司投影頁', () => {
  delete ssStub.tabs['光明燈管理'];
  const msg = run('setupLightSheet()');
  assert.ok(/主表「光明燈管理」就緒/.test(msg), msg);
  assert.ok(/只讀投影/.test(msg), msg);
  assert.deepStrictEqual(ssStub.created, ['光明燈管理', '聖文', '明典'], '應建立：' + ssStub.created);
});
t('公司分頁 = QUERY 投影主表，且主表資料不被碰', () => {
  run('setupLightSheet()');
  const sw = ssStub.tabs['聖文'], md = ssStub.tabs['明典'];
  assert.deepStrictEqual(sw._rows[0], HEADER, '投影頁也要有 A~H 表頭');
  assert.ok(/^=QUERY\('光明燈管理'!A2:H, "where upper\(A\) contains upper\("聖文"\)", 0\)$/.test(sw._formulas['2,1']), sw._formulas['2,1']);
  assert.ok(/upper\("明典"\)/.test(md._formulas['2,1']), md._formulas['2,1']);
  assert.strictEqual(sw._protected, true, '投影頁應設為編輯警告');
  assert.strictEqual(sheet._rows.length, 18, '主表資料不應被清掉');
  assert.strictEqual(Object.keys(sheet._formulas).length, 0, '主表不應被寫 QUERY 公式');
  assert.strictEqual(sw._rows[0].length, 8, '投影頁表頭應含備註欄共 8 欄');
});
t('其他家只有主表：不會冒出奇怪分頁', () => {
  run('setupLightSheet()');
  assert.strictEqual(Object.keys(ssStub.tabs).sort().join(','), '光明燈管理,明典,聖文');
  ai({ action: 'CREATE', details: D({ agent: '冠緯', temple: '新旺宮', spec: '5*7 OLED', total_count: 90 }) });
  post('小幫手 登記 冠緯 新旺宮 5*7 OLED 90盞');
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
  assert.ok(/冠緯/.test(sheet._rows[sheet._rows.length - 1][0]), '應寫進主表');
  assert.deepStrictEqual(ssStub.created, ['聖文', '明典'], '不應為冠緯另開分頁：' + ssStub.created);
});
t('SHEET_NAMES 拿掉兩行 -> 投影機制整個不執行', () => {
  run('CONFIG.SHEET_NAMES = { LIGHT_MGMT: "光明燈管理" }');
  const msg = run('setupLightSheet()');
  assert.ok(/未設定公司分頁/.test(msg), msg);
  assert.strictEqual(ssStub.created.length, 0, '不應建立任何分頁');
  run('CONFIG.SHEET_NAMES = { LIGHT_MGMT: "光明燈管理", SHENG_WEN: "聖文", MING_DIAN: "明典" }');
});
t('SPREADSHEET_ID 沒設 -> 明確報錯', () => {
  run('CONFIG.SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"');
  assert.throws(() => run('getSpreadsheet()'), /尚未設定 SPREADSHEET_ID/);
});


console.log('\n【新增時遇到同廟：A 承攬商納入判斷 + B 多列先列清單】');
const D2 = o => Object.assign({ agent: null, temple: null, spec: null, total_count: null, delivery_text: null, software: null, computer: null, remark: null }, o);
const create = (d, txt) => run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: d })}, ${JSON.stringify(txt || '')}, "U_DUP")`);
t('新廟名 -> 直接新增並標註第一筆', () => {
  const r = create(D2({ agent: '聖文', temple: '順安宮', spec: '5*7 OLED', total_count: 50 }));
  assert.ok(/新增成功/.test(r), r);
  assert.ok(/新廟名/.test(r), r);
});
t('同廟但規格不同 -> 仍新增，並列出該廟既有規格', () => {
  const r = create(D2({ agent: '聖文', temple: '金六結福德廟', spec: '7*9 OLED', total_count: 300 }));
  assert.ok(/新增成功/.test(r), r);
  assert.ok(/該廟既有規格：4\*5 OLED（4,472 盞）、5\*7 OLED（1,407 盞）/.test(r), r);
});
t('A) 同廟同規格但承攬商不同 -> 當作另一張單新增，並提示既有那筆', () => {
  const before = sheet._rows.length;
  const r = create(D2({ agent: '名典', temple: '石岡子乾元宮', spec: '5*7 OLED琥珀色', total_count: 300 }));
  assert.strictEqual(sheet._rows.length, before + 1, '應寫入：' + r);
  assert.ok(!/疑似重複登記/.test(r), r);
  assert.ok(/第 2 列已有同規格但承攬商為「聖文」/.test(r), r);
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][0], '名典');
});
t('同廟同規格且同承攬商 -> 擋下並給出正確的「修改」範例（用全名）', () => {
  const before = sheet._rows.length;
  const r = create(D2({ agent: '趴一', temple: '天成宮(中和)', spec: '4*5 OLED', total_count: 200 }));
  assert.strictEqual(sheet._rows.length, before, '不應寫入：' + r);
  assert.ok(/疑似重複登記】第 17 列/.test(r), r);
  assert.ok(/修改 天成宮\(中和\) 4\*5 OLED 總燈數變成 2500 盞/.test(r), r);
});
t('B) 只打「天成宮」兩館同規格 -> 不寫入、列出兩館、存補件草稿', () => {
  const before = sheet._rows.length;
  const r = create(D2({ agent: '聖文', temple: '天成宮', spec: '4*5 OLED', total_count: 200 }), '小幫手 天成宮 4*5 OLED 200盞');
  assert.strictEqual(sheet._rows.length, before, '不應寫入：' + r);
  assert.ok(/同名廟有 2 列相同規格/.test(r), r);
  assert.ok(/第 16 列：天成宮\(北投\)/.test(r), r);
  assert.ok(/第 17 列：天成宮\(中和\)/.test(r), r);
  assert.ok(run('!!loadDraft("U_DUP")'), '應留下補件草稿');
});
t('B 之後補全名 -> 接到草稿並指向正確那一館', () => {
  create(D2({ agent: '聖文', temple: '天成宮', spec: '4*5 OLED', total_count: 200 }), '小幫手 天成宮 4*5 OLED 200盞');
  const draft = run('loadDraft("U_DUP")');
  assert.ok(draft, '應有補件草稿');
  const merged = JSON.parse(run(`JSON.stringify(mergeDetails(${JSON.stringify(draft.details)}, ${JSON.stringify(D2({ temple: '天成宮(中和)', spec: '4*5 OLED', total_count: 200 }))}))`));
  assert.strictEqual(merged.temple, '天成宮(中和)');
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: merged })}, "天成宮(中和) 4*5 200盞", "U_DUP")`);
  // 草稿業務是預設「聖文」，既有那列是「趴一」→ 依 A 規則當作另一張單，但已指向正確的中和館
  assert.ok(/新增成功/.test(r), r);
  assert.ok(/第 17 列已有同規格但承攬商為「趴一」/.test(r), r);
  assert.ok(/天成宮\(中和\)/.test(sheet._rows[sheet._rows.length - 1][1]), JSON.stringify(sheet._rows[sheet._rows.length - 1]));
});
t('強制新增可以繞過一切比對', () => {
  const before = sheet._rows.length;
  const r = create(D2({ agent: '聖文', temple: '天成宮', spec: '4*5 OLED', total_count: 200 }), '強制新增');
  assert.strictEqual(sheet._rows.length, before + 1, r);
  assert.ok(/新增成功/.test(r), r);
});
t('合計列不會被當成同廟', () => {
  const r = create(D2({ agent: '聖文', temple: '合計燈數', spec: '4*5 OLED', total_count: 1 }));
  assert.ok(/新廟名/.test(r), r);
});

console.log('\n【LINE 送出層：5000 字上限與多則切分】');
t('短訊不變長、不切則', () => {
  const n = run(`prepareLineMessages("hello").length`);
  assert.strictEqual(n, 1);
});
t('超長內容依換行切則', () => {
  const big = Array.from({ length: 300 }, (_, i) => `第${i}行 光明燈登記說明 5*7 OLED琥珀色 2112盞 國10/17前`).join('\n');
  const chunks = JSON.parse(run(`JSON.stringify(prepareLineMessages(${JSON.stringify(big)}))`));
  assert.ok(chunks.length >= 2, '應切成多則：' + chunks.length);
  chunks.forEach(c => assert.ok(c.length <= 4800, '單則不超 4800：' + c.length));
});
t('切點太爛時硬切，且不超過上限', () => {
  const one = '甲'.repeat(9000);
  const chunks = JSON.parse(run(`JSON.stringify(splitLineText(${JSON.stringify(one)}, 4800))`));
  assert.strictEqual(chunks.length, 2);
  assert.strictEqual(chunks[0].length, 4800);
});
t('超過 5 則會省略並在末則說明', () => {
  const huge = Array.from({ length: 20 }, () => '乙'.repeat(4700) + '\n').join('');
  const chunks = JSON.parse(run(`JSON.stringify(prepareLineMessages(${JSON.stringify(huge)}))`));
  assert.strictEqual(chunks.length, 5, '最多 5 則');
  assert.ok(/內容過長，已省略/.test(chunks[4]), chunks[4].slice(-90));
});
t('LINE 非 200 時明確拋錯（不再靜靜不回訊）', () => {
  ctx.HTTP_CODE = 400;
  assert.throws(() => run(`sendLineReply("RT", "x")`), /LINE 送出失敗 HTTP 400/);
});

console.log('\n【聊天回覆字數上限】');
t('預設 600 字', () => assert.strictEqual(run('chatCharLimit()'), 600));
t('腳本屬性 CHAT_MAX_CHARS 可覆蓋，且夾在 80~2000', () => {
  run('CONFIG.CHAT.MAX_REPLY_CHARS = 300');
  assert.strictEqual(run('chatCharLimit()'), 300);
  run('CONFIG.CHAT.MAX_REPLY_CHARS = 5');
  assert.strictEqual(run('chatCharLimit()'), 80);
  run('CONFIG.CHAT.MAX_REPLY_CHARS = 99999');
  assert.strictEqual(run('chatCharLimit()'), 2000);
});
t('模型超長時程式再硬截一道', () => {
  run('CONFIG.CHAT.MAX_REPLY_CHARS = 120');
  const long = Array.from({ length: 20 }, (_, i) => `說明第${i}行，內容有點長用來測試截斷行為。`).join('\n');
  const clipped = run(`clipChatReply(${JSON.stringify(long)})`);
  assert.ok(clipped.length <= 200, '截斷後仍過長：' + clipped.length);
  assert.ok(/已限制 120 字/.test(clipped), clipped.slice(-60));
  assert.ok(!clipped.includes('說明第19行'), '尾段應被去掉');
});
t('空回覆不會回傳空字串', () => assert.strictEqual(run(`clipChatReply("   ")`), '（沒什麼回應，再問一次看看）'));
t('system prompt 裡寫明字數上限', () => assert.ok(/字以內（含標點）/.test(run(`chatSystemPrompt('操作說明','怎麼登記')`))));

console.log('\n【LLM 供應者路由（GEMINI / OLLAMA）】');
t('預設走 Gemini，不打 Ollama', () => {
  run(`askLLM({ system:'s', user:'u', wantJson:false, purpose:'chat' })`);
  assert.strictEqual(ollamaCalls.length, 0);
  assert.ok(geminiCalls.length >= 1, 'Gemini 應被呼叫');
});
t('CHAT_PROVIDER=OLLAMA 只讓聊天走 Ollama，解析仍 Gemini', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "test-key"');
  run(`askLLM({ system:'s', user:'你好', wantJson:false, purpose:'chat' })`);
  assert.strictEqual(ollamaCalls.length, 1, '聊天應打 Ollama');
  assert.strictEqual(ollamaCalls[0].model, 'gpt-oss:20b', '用預設模型名');
  run(`askLLM({ user:'登記 石岡子乾元宮 5*7 2112盞', wantJson:true, purpose:'parse' })`);
  assert.strictEqual(ollamaCalls.length, 1, '解析不應打 Ollama');
  assert.ok(geminiCalls.length >= 2, '解析仍走 Gemini');
});
t('Ollama 请求格式：messages 帶 system/history、stream=false', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "test-key"');
  run('CONFIG.LLM.OLLAMA_MODEL = "llama3.2:3b"');
  const r = run(`askLLM({ system:'你是小幫手', user:'怎麼改燈數', history:[{role:'user',text:'前一句'}], wantJson:false, purpose:'chat' })`);
  const body = ollamaCalls[ollamaCalls.length - 1];
  assert.strictEqual(body.model, 'llama3.2:3b');
  assert.strictEqual(body.stream, false);
  assert.strictEqual(body.messages[0].role, 'system');
  assert.strictEqual(body.messages[1].content, '前一句');
  assert.strictEqual(body.messages[2].content, '怎麼改燈數');
  assert.strictEqual(r, 'OK 從 Ollama 來的回覆');
});
t('要 Ollama 但沒設 key -> 明確告訴去哪設', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  assert.throws(() => run(`askLLM({ system:'s', user:'u', purpose:'chat' })`), /OLLAMA_API_KEY/);
});
t('wantJson 時 Ollama 加 format=json 並解析', () => {
  run('CONFIG.LLM.PROVIDER = "OLLAMA"; CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "k"');
  ctx.OLLAMA_REPLY = '```json\n{"action":"READ","details":{}}\n```';
  const j = run(`askLLM({ system:'s', user:'u', wantJson:true, purpose:'parse' })`);
  assert.strictEqual(ollamaCalls[ollamaCalls.length - 1].format, 'json');
  assert.strictEqual(j.action, 'READ');
});
t('LLM_PROVIDER=OLLAMA 但回出爛 JSON -> 建議退回 GEMINI 的錯誤訊息', () => {
  run('CONFIG.LLM.PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "k"');
  ctx.OLLAMA_REPLY = '我不知道你在說什麼';
  assert.throws(() => run(`askLLM({ system:'s', user:'u', wantJson:true, purpose:'parse' })`), /建議把解析改回 GEMINI/);
  run('CONFIG.LLM.PROVIDER = "GEMINI"');
});

console.log('\n【聊天模式指令與流程】');
t('聊天指令辨識', () => {
  assert.strictEqual(run(`parseChatCommand("聊天").action`), 'on');
  assert.strictEqual(run(`parseChatCommand("/chat").action`), 'on');
  assert.strictEqual(run(`parseChatCommand("小幫手 聊天".replace("小幫手 ","")).action`), 'on');
  assert.strictEqual(run(`parseChatCommand("結束聊天").action`), 'off');
  assert.strictEqual(run(`parseChatCommand("/chat off").action`), 'off');
  assert.strictEqual(run(`parseChatCommand("角色").action`), 'roles');
  assert.strictEqual(run(`parseChatCommand("角色 文案助手").role`), '文案助手');
  assert.strictEqual(run(`parseChatCommand("登記 石岡子乾元宮 5*7 2112盞")`), null, '登記句不應被當成聊天指令');
});
t('進入/結束聊天狀態與過期自動關', () => {
  run(`setChatState("U_C1", true)`);
  assert.ok(run(`!!getChatState("U_C1")`), '應該在聊天中');
  assert.ok(run(`handleChatCommand("U_C1", {action:'off'})`).indexOf('已結束聊天') !== -1);
  assert.strictEqual(run(`getChatState("U_C1")`), null);
  run(`setChatState("U_C2", true)`);
  run(`(function(){ var m=${'{}'}; var s=readChatStates(); s.U_C2.at = Date.now() - (CONFIG.CHAT.MODE_TTL*1000 + 5000); })()`);
  run(`(function(){ var m=readChatStates(); m.U_C2.at = Date.now() - (CONFIG.CHAT.MODE_TTL*1000 + 5000); writeChatStates(m); })()`);
  assert.strictEqual(run(`getChatState("U_C2")`), null, '超過 60 分鐘沒動作應自動關閉');
});
t('角色清單與切換、不存在的角色', () => {
  const list = run(`handleChatCommand("U_C3", {action:'roles'})`);
  assert.ok(/操作說明 ← 使用中/.test(list), list);
  assert.ok(/禮俗顧問|文案助手/.test(list), list);
  assert.ok(/已切換角色：文案助手/.test(run(`handleChatCommand("U_C3", {action:'role', role:'文案助手'})`)));
  assert.ok(/沒有這個角色/.test(run(`handleChatCommand("U_C3", {action:'role', role:'外星人'})`)));
});
t('聊天模式下：未帶喚醒字的問句也會被處理', () => {
  run(`setChatState("U_TEST", true)`);
  assert.strictEqual(run(`decideFlow("送燈時間要怎麼寫？", false, !!getChatState("U_TEST")).mode`), 'ai');
  run(`setChatState("U_TEST", false)`);
  assert.strictEqual(run(`decideFlow("送燈時間要怎麼寫？", false, false).mode`), 'silent');
});
t('聊天模式 e2e：純問句直接聊，回覆帶角色前綴與字數限制', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "k"');
  ctx.OLLAMA_REPLY = '登记时把庙名、规格、盏数写在「小幫手」后面即可。';
  run(`setChatState("U_TEST", true)`);
  post('那要怎麼打？');
  assert.ok(/^💬【操作說明】/.test(sent[0]), sent[0]);
  assert.ok(/登记时把庙名/.test(sent[0]), sent[0]);
  assert.strictEqual(ollamaCalls.length, 1, '只該打 Ollama 一次');
});
t('聊天模式中傳登記句 -> 仍優先入表（不被當成聊天）', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "k"');
  run(`setChatState("U_TEST", true)`);
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '永照宮', spec: '5*7 OLED', total_count: 66 }) });
  const before = sheet._rows.length;
  post('小幫手 登記 永照宮 5*7 OLED 66盞');
  assert.strictEqual(sheet._rows.length, before + 1, '應寫入主表');
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
});
t('對話記憶只留最近 5 輪', () => {
  for (let i = 1; i <= 9; i++) run(`pushHistory("U_H", "問${i}", "答${i}")`);
  const h = JSON.parse(run(`JSON.stringify(loadHistory("U_H"))`));
  assert.strictEqual(h.length, 10, '5 輪 = 10 則');
  assert.strictEqual(h[0].text, '問5');
  assert.strictEqual(h[9].text, '答9');
});
t('下一句聊天會帶著前幾輪歷史送出去', () => {
  run('CONFIG.LLM.CHAT_PROVIDER = "OLLAMA"');
  run('CONFIG.OLLAMA_API_KEY = "k"');
  run(`setChatState("U_TEST2", true)`);
  ctx.OLLAMA_REPLY = '第一答';
  run(`chatAnswer("U_TEST2", "第一問")`);
  ctx.OLLAMA_REPLY = '第二答';
  run(`chatAnswer("U_TEST2", "第二問")`);
  const body = ollamaCalls[ollamaCalls.length - 1];
  const flat = JSON.stringify(body.messages);
  assert.ok(flat.indexOf('第一問') !== -1 && flat.indexOf('第一答') !== -1, '應帶上下文：' + flat.slice(0, 300));
});

console.log('\n【操作手冊來源自動偵測】');
t('沒文件沒分頁 -> 退回 /help 內建文案', () => {
  const m = run(`chatSystemPrompt('操作說明', '怎麼登記')`);
  assert.ok(/系統內建說明/.test(m) || /使用方式/.test(m), m.slice(0, 400));
});
t('有「操作手冊」分頁 -> 讀成條目並塞進 prompt', () => {
  ssStub.tabs['操作手冊'] = makeSheet([
    ['主題', '內容', '範例句'],
    ['送燈時間怎麼寫', '國曆寫「國10/17前」；農曆系統不敢換算', '小幫手 石岡子乾元宮 5*7 2112盞 國10/17前'],
    ['改燈數', '只講要改的欄位，其他不會清空', '小幫手 修改 新化武廟 4*5 OLED 總燈數變成5300盞'],
    ['合計燈數', '', '']
  ], '操作手冊');
  const items = JSON.parse(run(`JSON.stringify(manualItems())`));
  assert.strictEqual(items.length, 2, '應跳過合計列：' + JSON.stringify(items));
  const prompt = run(`chatSystemPrompt('操作說明', '送燈時間要怎麼寫')`);
  assert.ok(/【送燈時間怎麼寫】/.test(prompt), prompt.slice(-500));
  assert.ok(/農曆系統不敢換算/.test(prompt));
});
t('手冊很長時只挑相關條目（省 token）', () => {
  const rows = [['主題', '內容', '範例句']];
  ['軟體', '電腦', '規格', '備註', '送燈時間', '廟宇別名', '強制新增', '補件'].forEach((k, i) => rows.push([k, k + '的說明內容，第' + i + '段', '']));
  ssStub.tabs['操作手冊'] = makeSheet(rows, '操作手冊');
  run(`CacheService.getScriptCache().remove('CHATMANUAL')`);
  const picked = run(`pickManual('送燈時間 要怎麼寫')`);
  assert.ok(/【送燈時間】/.test(picked), picked);
  assert.ok(picked.indexOf('【軟體】') === -1, '無關條目不應進 prompt：' + picked);
});


console.log('\n【操作手冊：Drive 資料夾來源與可用格式】');
const setFolder = files => { ctx.FOLDER = files; run(`CacheService.getScriptCache().remove('CHATMANUAL')`) };
t('Google 文件：段落切成條目，主題帶檔名方便引用', () => {
  setFolder([{ id: 'D1', name: '光明燈操作手冊', mime: 'application/vnd.google-apps.document',
    text: '送燈時間怎麼寫\n國曆寫「國10/17前」，農曆系統不敢換算。\n\n改燈數\n只講要改的欄位，其他不會清空。' }]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"');
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items.length, 2, JSON.stringify(items));
  assert.strictEqual(items[0].topic, '光明燈操作手冊｜送燈時間怎麼寫');
  assert.ok(/農曆系統不敢換算/.test(items[0].body));
});
t('.md 與 .txt 也能讀，並清掉 markdown 符號', () => {
  setFolder([{ id: 'T1', name: 'manual.md', mime: 'text/markdown', text: '## 強制新增\n同廟同規格要當**第二批**時加這四個字。\n\n## 備註欄\n寫「備註：…」才會進 H 欄。' }]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"');
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items.length, 2);
  assert.ok(!/#|\*\*/.test(items[0].topic), '主題不該殘留 # 或 **：' + items[0].topic);
  assert.ok(items[0].body.indexOf('**') === -1, items[0].body);
});
t('PDF 與上傳的 .docx 會被跳過（讀不到文字）', () => {
  setFolder([
    { id: 'P1', name: '說明.pdf', mime: 'application/pdf', text: '' },
    { id: 'P2', name: '說明.docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', text: '' },
    { id: 'P3', name: '真的可用', mime: 'application/vnd.google-apps.document', text: '只有這筆可用' }
  ]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"');
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items.length, 1, JSON.stringify(items));
  assert.strictEqual(items[0].topic, '真的可用｜只有這筆可用');
});
t('MANUAL_MAX_FILES 限制讀取的檔案數', () => {
  setFolder([
    { id: 'A', name: 'A', mime: 'application/vnd.google-apps.document', text: '第一份\n內容A' },
    { id: 'B', name: 'B', mime: 'application/vnd.google-apps.document', text: '第二份\n內容B' }
  ]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"; CONFIG.CHAT.MANUAL_MAX_FILES = 1');
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items.length, 1, JSON.stringify(items));
  assert.ok(/內容A|第一份/.test(items[0].topic + items[0].body));
});
t('檔名含「合計」的文件整份略過', () => {
  setFolder([{ id: 'X', name: '合計表', mime: 'application/vnd.google-apps.document', text: '不要讀我\n內容' }]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"');
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items[0].topic, '系統內建說明', '該文件應被略過，退回 /help 内建文案：' + JSON.stringify(items.map(x => x.topic)));
});
t('沒設資料夾 ID -> 完全不碰 Drive', () => {
  setFolder([{ id: 'D1', name: '光明燈操作手冊', mime: 'application/vnd.google-apps.document', text: '送燈時間\n內容' }]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = ""');
  const prompt = run(`chatSystemPrompt('操作說明', '怎麼登記')`);
  assert.ok(/使用方式|內建說明/.test(prompt), prompt.slice(0, 200));
});
t('優先序：單檔 Doc ID > 資料夾 > 分頁', () => {
  setFolder([{ id: 'D1', name: '資料夾文件', mime: 'application/vnd.google-apps.document', text: '資料夾主題\n內容' }]);
  run('CONFIG.CHAT.MANUAL_FOLDER_ID = "F1"; CONFIG.CHAT.MANUAL_DOC_ID = "S1"');
  ctx.FOLDER.push({ id: 'S1', name: '單一文件', mime: 'application/vnd.google-apps.document', text: '單檔主題\n內容', standalone: true });
  const items = JSON.parse(run('JSON.stringify(manualItems())'));
  assert.strictEqual(items[0].topic, '單檔主題', JSON.stringify(items));
});

console.log(`\n結果：${pass} 通過 / ${fail} 失敗`);
process.exit(fail ? 1 : 0);
