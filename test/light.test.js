const fs = require('fs'), vm = require('vm'), assert = require('assert');
const SRC = require('path').join(__dirname, '..', 'light.gs');

const HEADER = ['業務', '廟宇', '規格', '總燈數', '送燈日期', '軟體', '電腦', '正規日期', '備註'];
const SEED = [
  HEADER,
  ['聖文', '石岡子乾元宮', '5*7 OLED琥珀色', 2112, '已送燈', '', '研華', '', '去年已換 5*7 全彩'],
  ['聖文', '桃園廣盛壇', '5*7 OLED琥珀色', 456, '國10/17前', '其他', '廟', ''],
  ['聖文', '桃園廣盛壇', '7*9 OLED琥珀色', 576, '', '其他', '廟', ''],
  ['甫穎', '仁武保安宮', '5*7 OLED琥珀色', 2475, '國10/22', '甫穎', '廟', ''],
  ['聖文', '田洋城隍廟', '5*7 OLED琥珀色', 1600, '國11/07前', '廟幫手', '廟', ''],
  ['士豪', '新化武廟', '4*5 OLED琥珀色', 5238, '國11/15前', '冠緯', '研華', ''],
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
let cacheStore = {}, sent = [];
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
      if (String(url).indexOf('/v2/bot/profile/') !== -1) {
        return { getResponseCode: () => (ctx.PROFILE_NAME ? 200 : 404), getContentText: () => JSON.stringify({ displayName: ctx.PROFILE_NAME || '' }) };
      }
      sent.push(JSON.parse(opt.payload).messages[0].text);
      return { getResponseCode: () => 200, getContentText: () => '{}' };
    } },
    HtmlService: { createHtmlOutput: x => ({ content: x }) },
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
  run('CONFIG.DEFAULT_AGENT = "聖文"; CONFIG.AGENT_FROM_LINE_PROFILE = false');
  cacheStore = {}; sent = [];
  ctx.__props.REPLY_MODE = ''; delete ctx.__props.USER_IDENTS; run('CONFIG.AGENT_FROM_LINE_PROFILE = false');
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
  assert.strictEqual(JSON.stringify(last), JSON.stringify(['聖文', '石岡子乾元宮', '5*7 OLED琥珀色', 300, '國11/20前', '廟幫手', '研華', '2026-11-20', '']));
});
t('一般新增寫入 8 欄', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '冠宇', temple: '行天宮', spec: '5*7 OLED琥珀色', total_count: 880, delivery_text: '國115/02/09前', software: '冠宇', computer: '研華' } })}, "x")`);
  assert.ok(/新增成功/.test(r), r);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[3], 880); assert.strictEqual(last[7], '2026-02-09');
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
  assert.strictEqual(row[7], '2026-03-01', '輔助日期應同步：' + row[7]);
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
t('右側有公式欄 -> 新增只寫 A~I 九欄', () => {
  sheet._rows = SEED.map(r => r.concat(['小計', 1]));
  run(`handleDataRouting(${JSON.stringify({ action: 'CREATE', details: { agent: '冠宇', temple: '寶林宮', spec: '5*7 OLED', total_count: 66, delivery_text: '國11/30前', software: null, computer: null } })}, "x")`);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last.length, 9, '應只有 A~I 九欄：' + JSON.stringify(last));
  assert.strictEqual(last[7], '2026-11-30');
});
t('右側有公式欄 -> 修改仍只動指定欄', () => {
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
  ctx.__props.REPLY_MODE = 'ALL';
  assert.strictEqual(flow('查一下 天成宮 有哪些燈'), 'ai');
  assert.strictEqual(run('shouldReplyFor("READ")'), true);
});
t('REPLY_MODE=off 不影響（非 ALL 值）', () => {
  ctx.__props.REPLY_MODE = 'WRITE';
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
  ctx.__props.REPLY_MODE = 'ALL';
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
  assert.ok(/共 4 列/.test(r), r);
  assert.ok(/第 2 列：聖文 → 聖文企業（石岡子乾元宮）/.test(r), r);
});
t('真的改名 -> 4 列更新，合計列不碰', () => {
  const r = run(`renameAgent("聖文", "聖文企業")`);
  assert.ok(/已完成/.test(r), r);
  [1, 2, 3, 5].forEach(i => assert.strictEqual(sheet._rows[i][0], '聖文企業', '第 ' + (i + 1) + ' 列'));
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

console.log('\n【備註欄與曆別（國曆/民國/農曆）】');
t('明示備註才會寫入 I 欄', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '永樂宮', spec: '5*7 OLED', total_count: 88, remark: '分兩批送，第二批國12/01前' }) });
  post('小幫手 登記 永樂宮 5*7 OLED 88盞 備註：分兩批送，第二批國12/01前');
  assert.ok(/新增成功/.test(sent[0]), sent[0]);
  const last = sheet._rows[sheet._rows.length - 1];
  assert.strictEqual(last[8], '分兩批送，第二批國12/01前', '整列=' + JSON.stringify(last) + '｜回覆=' + sent[0]);
  assert.strictEqual(last.length, 9);
});
t('沒提備註 -> I 欄留空，不拿整句原文亂填', () => {
  ai({ action: 'CREATE', details: D({ agent: '聖文', temple: '樂善宮', spec: '5*7 OLED', total_count: 10 }) });
  post('小幫手 登記 樂善宮 5*7 OLED 10盞');
  assert.strictEqual(sheet._rows[sheet._rows.length - 1][8], '');
});
t('修改可單獨改備註', () => {
  const r = run(`handleDataRouting(${JSON.stringify({ action: 'UPDATE', details: D({ temple: '石岡子乾元宮', spec: '5*7', remark: '已改全彩' }) })}, "x")`);
  assert.ok(/備註：去年已換 5\*7 全彩 → 已改全彩/.test(r), r);
  assert.strictEqual(sheet._rows[1][8], '已改全彩');
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
  assert.strictEqual(last[7], '', 'H 欄不應亂推：' + last[7]);
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
  assert.ok(/^=QUERY\('光明燈管理'!A2:I, "where upper\(A\) contains upper\("聖文"\)", 0\)$/.test(sw._formulas['2,1']), sw._formulas['2,1']);
  assert.ok(/upper\("明典"\)/.test(md._formulas['2,1']), md._formulas['2,1']);
  assert.strictEqual(sw._protected, true, '投影頁應設為編輯警告');
  assert.strictEqual(sheet._rows.length, 18, '主表資料不應被清掉');
  assert.strictEqual(Object.keys(sheet._formulas).length, 0, '主表不應被寫 QUERY 公式');
  assert.strictEqual(sw._rows[0].length, 9, '投影頁表頭應含備註欄共 9 欄');
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

console.log(`\n結果：${pass} 通過 / ${fail} 失敗`);
process.exit(fail ? 1 : 0);
