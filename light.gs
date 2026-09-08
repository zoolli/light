// ==================== 1. 光明燈規格管理設定區 ====================
// 所有可调参数都在这里，一律「腳本屬性有值就用、沒值用預設」。
// 函式宣告會被提前解析，所以 CONFIG 裡可以直接呼叫 sp()。
// 注意：金鑰只寫佔位字串，實際值請填在 GAS「專案設定 → 腳本屬性」，絕對不要寫進本檔（已推到 GitHub）
function sp(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    if (v === null || String(v).trim() === '') return fallback;
    return String(v).trim();
  } catch (e) { return fallback; }
}

function spInt(key, fallback) {
  var n = parseInt(sp(key, ''), 10);
  return isNaN(n) ? fallback : n;
}

function spList(key, fallback) {
  return String(sp(key, fallback)).split(/[,，]/).map(function (x) { return x.trim(); }).filter(function (x) { return x; });
}

function spOn(key, fallback) {
  var v = sp(key, fallback ? 'ON' : '');
  return /^(ON|YES|1|TRUE|ALL)$/i.test(v);
}

var CONFIG = {
  // ---- 金鑰與資料位置 ----
  LINE_ACCESS_TOKEN: sp('LINE_ACCESS_TOKEN', 'YOUR_LINE_ACCESS_TOKEN'),
  GEMINI_API_KEY: sp('GEMINI_API_KEY', 'YOUR_GEMINI_API_KEY'),
  OLLAMA_API_KEY: sp('OLLAMA_API_KEY', 'YOUR_OLLAMA_API_KEY'),      // ollama.com/settings/keys 產生
  GEMINI_MODEL: sp('GEMINI_MODEL', 'gemini-3.6-flash'),
  SPREADSHEET_ID: sp('SPREADSHEET_ID', (function () {
    try { return SpreadsheetApp.getActiveSpreadsheet().getId(); } catch (e) { return 'YOUR_SPREADSHEET_ID'; }
  })()),

  // ---- 分頁名稱：LIGHT_MGMT 是主表（全部資料）；其餘是主表的唯讀投影 ----
  SHEET_NAMES: {
    LIGHT_MGMT: sp('SHEET_NAME', '光明燈管理'),
    SHENG_WEN: '聖文',
    MING_DIAN: '明典'
  },

  // ---- 欄位索引（1 起算）。業務與廟宇分開兩欄，其他資訊收進備註 ----
  COLUMNS: {
    AGENT: 1, TEMPLE: 2, SPEC: 3, TOTAL: 4, DELIVERY: 5,
    SOFTWARE: 6, COMPUTER: 7, REMARK: 8
  },
  // 系統推得出的國曆日期以〔國YYYY-MM-DD〕標記併入備註欄，沒有獨立的正規日期欄
  HEADERS: ['業務', '廟宇', '規格', '總燈數', '送燈時間', '軟體', '電腦', '備註'],
  REMARK_DATE_TAG: !/^OFF|NO|0|FALSE$/i.test(sp('DATE_TAG_IN_REMARK', 'ON')),

  // ---- 回覆與指令行為 ----
  WAKE_WORDS: spList('WAKE_WORDS', '小幫手,小帮,幫手,助理'),
  REPLY_MODE: sp('REPLY_MODE', ''),                 // 設 ALL = 查詢與無關訊息也回覆
  REQUIRED_FIELDS: ['agent', 'temple', 'spec', 'total_count'],
  FIELD_LABELS: {
    agent: '業務／公司名', temple: '廟宇', spec: '規格', total_count: '總燈數',
    delivery_text: '送燈時間', software: '軟體', computer: '電腦', remark: '備註'
  },
  DEFAULT_AGENT: sp('DEFAULT_AGENT', '聖文'),        // 沒寫業務時的預設承攬商；設空白則回問
  AGENT_FROM_LINE_PROFILE: spOn('AGENT_FROM_LINE_PROFILE', false),  // 預設 OFF：暱稱常是個人綽號
  DRAFT_TTL: spInt('DRAFT_TTL', 600),               // 補件草稿保留秒數（上限 600）
  MAX_LIST_ROWS: spInt('MAX_LIST_ROWS', 12),

  // ---- LLM 供應者：GEMINI（免錢、預設）或 OLLAMA（免費額度，1 併發）----
  LLM: {
    PROVIDER: sp('LLM_PROVIDER', 'GEMINI'),         // 登記解析用哪家
    CHAT_PROVIDER: sp('CHAT_PROVIDER', ''),         // 留空=跟 PROVIDER；可單獨給聊天用 Ollama
    OLLAMA_BASE_URL: sp('OLLAMA_BASE_URL', 'https://ollama.com'),
    OLLAMA_MODEL: sp('OLLAMA_MODEL', 'gpt-oss:20b'),// 免費層建議 level 1~2 輕量模型
    OLLAMA_TIMEOUT_HINT: 'Ollama 免費層只有 1 個併發，兩人同時發言會排隊或被拒（429）；' +
                         '要穩就把 LLM_PROVIDER 設回 GEMINI，或只在 CHAT_PROVIDER 用 OLLAMA'
  },

  // ---- 聊天模式（小幫手 聊天／結束聊天、/角色）----
  CHAT: {
    DEFAULT_ROLE: sp('CHAT_DEFAULT_ROLE', '操作說明'),
    HISTORY_TURNS: spInt('CHAT_HISTORY_TURNS', 5),  // 帶最近幾輪
    HISTORY_TTL: spInt('CHAT_HISTORY_TTL', 600),
    MODE_TTL: spInt('CHAT_MODE_TTL', 3600),         // 60 分鐘沒動作自動關回沉默模式
    MAX_REPLY_CHARS: spInt('CHAT_MAX_CHARS', 600),  // 回覆字數上限（含標點）
    MANUAL_SHEET: sp('CHAT_MANUAL_SHEET', '操作手冊'),
    MANUAL_DOC_ID: sp('CHAT_MANUAL_DOC_ID', ''),    // 設了就會走 Google 文件（會多要求 Docs 授權）
    MANUAL_MAX_CHARS: spInt('CHAT_MANUAL_MAX_CHARS', 4000),
    ROLES: {
      '操作說明': '你是「宮廟光明燈管理系統」的操作小幫手，專門教導經辦人如何用 LINE 登記、修改光明燈規格。' +
                  '只根據提供的操作手冊與系統規則回答；手冊沒寫的不要編，直接說「這個我不確定，請問管理員」。' +
                  '先給結論，再條列重點，最後附一行可直接複製的輸入範例。',
      '禮俗顧問': '你是宮廟光明燈相關禮俗的說明人員，用繁體中文簡短回答太歲燈、光明燈、安奉與送燈等習俗問題；' +
                  '各家廟方做法不同時要說明「以貴廟慣例為準」，不斷言對錯，也不涉及勸信。',
      '文案助手': '你是燈務文案助手，依要求產出簡短、繁體中文、適合 LINE 或紅單列印的公告、提醒與說明文字，語氣禮貌樸實。',
      '通用助理': '你是這個系統內的一般問答助理，以繁體中文簡短回答與工作相關的問題；不確定就明說不確定。'
    }
  }
};

// ==================== 2. LINE Webhook 接收端（網頁回應優化版） ====================
function doPost(e) {
  try {
    // 防禦機制：若完全無資料傳入（例如直接對該網址發送空 POST），回傳一個簡易網頁畫面
    if (!e || !e.postData || !e.postData.contents) {
      return HtmlService.createHtmlOutput("<html><body style='font-family:sans-serif; text-align:center; padding-top:50px;'><h2>⛩️ 光明燈系統後端 Webhook</h2><p style='color:green;'>API 介面運作正常，請由 LINE 端傳送資料。</p></body></html>");
    }

    // 解析 LINE 傳來的 JSON 資料
    var jsonData = JSON.parse(e.postData.contents);
    var event = jsonData.events;
    
    // LINE 後台點選「Verify」驗證時的空測試包防禦
    if (!event || event.length === 0) {
      return HtmlService.createHtmlOutput("<html><body><h1>LINE Webhook Verified Successfully!</h1></body></html>");
    }
    
    var firstEvent = event[0]; // 取出第一筆事件
    
    // 檢查是否為文字訊息，若符合才啟動 Gemini AI 與試算表寫入
    if (firstEvent && firstEvent.replyToken && firstEvent.type === 'message' && firstEvent.message.type === 'text') {
      var replyToken = firstEvent.replyToken;
      var userMessage = firstEvent.message.text;
      var userId = (firstEvent.source && firstEvent.source.userId) || 'unknown-user';
      var draft = loadDraft(userId);
      var chatState = getChatState(userId);

      // 回覆策略：需要帶喚醒字（或在補件期間）才處理；查詢與不相關訊息保持沉默
      var flow = decideFlow(userMessage, !!draft, !!chatState);

      if (flow.mode === 'silent') {
        console.log('🔇 不回覆（' + flow.reason + '）：' + userMessage);
      } else if (flow.mode === 'command') {
        sendLineReply(replyToken, flow.handler === 'help' ? handleHelpCommand() : handleModelCommand(userMessage));
      } else if (flow.mode === 'chatcmd') {
        sendLineReply(replyToken, handleChatCommand(userId, flow.cmd));
      } else if (flow.mode === 'empty') {
        sendLineReply(replyToken, draft ? askMissingMessage(missingRequired(draft.details), draft.details) : handleHelpCommand());
      } else {
        var bodyText = flow.text;
        var idtCmd = parseIdentityCommand(bodyText);

        if (idtCmd) {
          clearDraft(userId);
          var label = bindIdentity(userId, idtCmd);
          sendLineReply(replyToken, '✅ 已記住您的業務身分：「' + label + '」\n之後登記不用每次都寫業務，未填時我就自動帶入。\n・公司派來的窗口：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 我公司 亞盛燈業 我叫 李小華 → 顯示「亞盛燈業-李小華」\n・換公司／離職：再傳一次即可蓋掉（例：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 我是 聖文）');
        } else if (flow.chat && !draft && !looksLikeLightData(bodyText)) {
          // 聊天模式中的純問句：直接聊，不必先花一次登記解析的呼叫
          sendLineReply(replyToken, chatAnswer(userId, bodyText));
        } else {
          var aiInput = draft ? (draft.raw + '；補充：' + bodyText) : bodyText;
          var aiResult = analyzeMessageWithGemini(aiInput);
          var act = String((aiResult && aiResult.action) || '').toUpperCase();

          if (act === 'CREATE') {
            var merged = mergeDetails(draft ? draft.details : null, aiResult.details);
            var agentNote = '';
            if (String(merged.agent || '').trim() === '') {
              var idt = getUserIdentity(userId);
              if (idt.source === 'bind' && idt.name) {
                merged.agent = idt.name;
                agentNote = '\n👤 業務自動帶入（您的綁定）：' + idt.name + '（換人/換公司：傳「我是 ○○」或「我公司 ○○ 我叫 ○○」）';
              } else if (CONFIG.DEFAULT_AGENT) {
                merged.agent = CONFIG.DEFAULT_AGENT;
                agentNote = '\n👤 業務預設帶入：' + CONFIG.DEFAULT_AGENT + '（若是别家訂單請寫在句首，或傳「我是 ○○」綁定您的公司）';
              } else if (idt.name) {
                merged.agent = idt.name;
                agentNote = '\n👤 業務自動帶入 LINE 暱稱「' + idt.name + '」，若應填公司名稱請傳「我公司 您的公司名」修正。';
              }
            }
            var miss = missingRequired(merged);
            if (miss.length > 0 && !/直接新增|強制新增|先佔位/.test(bodyText)) {
              saveDraft(userId, { raw: draft ? (draft.raw + '；' + bodyText) : bodyText, details: merged });
              sendLineReply(replyToken, askMissingMessage(miss, merged));
            } else {
              clearDraft(userId);
              sendLineReply(replyToken, handleDataRouting({ action: 'CREATE', details: merged }, aiInput, userId) + agentNote);
            }
          } else if (shouldReplyFor(act)) {
            if (draft) clearDraft(userId);
            sendLineReply(replyToken, handleDataRouting(aiResult, aiInput, userId));
          } else if (flow.chat) {
            // 聊天模式下被判定為查詢／無關 → 当成對話回，不沉默
            sendLineReply(replyToken, chatAnswer(userId, bodyText));
          } else {
            console.log('🔇 不回覆（AI 判定意圖 ' + (act || '未知') + '，非新增/修改）：' + bodyText);
          }
        }
      }
    }

  } catch (error) {
    console.error('doPost 發生錯誤: ' + error.toString());
    // 只有「使用者本來就期望有回應」的訊息（指令或寫入類）才打擾；沉默類只留日誌
    try {
      var errEvent = JSON.parse(e.postData.contents).events[0];
      var errText = (errEvent && errEvent.message && errEvent.message.text) || '';
      var errUid = (errEvent && errEvent.source && errEvent.source.userId) || 'unknown-user';
      if (errEvent && errEvent.replyToken && decideFlow(errText, !!loadDraft(errUid), !!getChatState(errUid)).mode !== 'silent') {
        sendLineReply(errEvent.replyToken, "【⚠️ 系統異常】" + error.toString());
      }
    } catch (e2) {}
  }
  
  // 最終一定要回傳 HtmlOutput，確保 LINE 伺服器拿到 HTTP 200 成功狀態代碼
  return HtmlService.createHtmlOutput("OK");
}

// ==================== 2.3 喚醒字與補件草稿 ====================
// 回傳「去掉喚醒字後的內容」；完全找不到喚醒字則回傳 null
function stripWakeWord(text) {
  var t = String(text || '').trim();
  var words = (CONFIG.WAKE_WORDS || []).slice().sort(function(a, b) { return String(b).length - String(a).length; });
  var i, w, cut = /^[\s,，。、:：；;！!？?～~-]+/;
  for (i = 0; i < words.length; i++) {
    w = String(words[i] || '').trim();
    if (w && t.substring(0, w.length) === w) return t.substring(w.length).replace(cut, '').trim();
  }
  for (i = 0; i < words.length; i++) {
    w = String(words[i] || '').trim();
    var pos = w ? t.indexOf(w) : -1;
    if (pos > 0 && pos <= 8) return (t.substring(0, pos) + ' ' + t.substring(pos + w.length)).replace(cut, '').trim();
  }
  return null;
}

function hasWakeWord(text) {
  return stripWakeWord(text) !== null;
}

function draftKey(userId) { return 'LIGHTDRAFT_' + userId; }

function loadDraft(userId) {
  try {
    var c = CacheService.getScriptCache().get(draftKey(userId));
    if (c) return JSON.parse(c);
  } catch (e) {}
  return null;
}

function saveDraft(userId, obj) {
  try { CacheService.getScriptCache().put(draftKey(userId), JSON.stringify(obj), CONFIG.DRAFT_TTL); } catch (e) {}
}

function clearDraft(userId) {
  try { CacheService.getScriptCache().remove(draftKey(userId)); } catch (e) {}
}

// 新取得的欄位覆蓋舊草稿（空格不覆蓋）
// 合併草稿與新取得的欄位（新值優先，空值不覆蓋）；欄位清單直接看 FIELD_LABELS，新增欄位不用改這裡
function mergeDetails(base, extra) {
  var out = {}, keys = Object.keys(CONFIG.FIELD_LABELS);
  keys.forEach(function(k) {
    var b = base ? base[k] : null, x = extra ? extra[k] : null;
    out[k] = (x === null || x === undefined || String(x).trim() === '') ? (b === undefined ? null : b) : x;
  });
  return out;
}

function missingRequired(details) {
  var miss = [];
  (CONFIG.REQUIRED_FIELDS || []).forEach(function(k) {
    var v = details ? details[k] : null;
    if (v === null || v === undefined || String(v).trim() === '') miss.push(k);
  });
  return miss;
}

var FIELD_PLACEHOLDERS = {
  agent: '聖文', temple: '石岡子乾元宮', spec: '5*7 OLED琥珀色', total_count: '2112盞',
  delivery_text: '國10/17前', software: '廟幫手', computer: '研華'
};

function askMissingMessage(miss, details) {
  var L = CONFIG.FIELD_LABELS;
  var got = [];
  Object.keys(L).forEach(function(k) {
    var v = details ? details[k] : null;
    if (v !== null && v !== undefined && String(v).trim() !== '') got.push(L[k] + ' ' + v);
  });
  var example = (CONFIG.WAKE_WORDS[0] || '小幫手') + ' ' +
    miss.map(function(k) { return FIELD_PLACEHOLDERS[k] || L[k]; }).join(' ');
  var lines = [
    '【🟡 資料還沒齊全，暫未登記】',
    '📥 已收到：' + (got.length ? got.join('、') : '（無）'),
    '❓ 還需要：' + miss.map(function(k) { return L[k]; }).join('、'),
    '',
    '✍️ 可以直接補：' + example
  ];
  if (miss.indexOf('agent') !== -1) lines.push('　（業務只需綁定一次：「我是 聖文」或「我公司 亞盛燈業 我叫 李小華」）');
  lines.push('⏳ ' + Math.round(CONFIG.DRAFT_TTL / 60) + ' 分鐘內補件我會自動合併，不必重打整串；' +
             '若想先佔位之後再補，請加「直接新增」');
  return lines.join('\n');
}

// ==================== 2.35 業務身分綁定（公司／派來的人員）與 LINE 姓名 ====================
// 業務欄可能填：自家業務名、公司名、或公司派來的窗口。因此身分拆成 company + person 兩段，
// 寫進試算表時顯示為「公司-人員」（只有一段時就只顯示那一段）。
var IDENTITY_KEY = 'USER_IDENTS';
var PERSON_KEYS = { '窗口': 1, '人員': 1, '我叫': 1, '名字叫': 1, '姓名': 1, '經辦人': 1 };

function readIdentities() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(IDENTITY_KEY) || '{}'); } catch (e) { return {}; }
}

function identityLabel(idt) {
  if (!idt) return '';
  var c = String(idt.company || '').trim(), pe = String(idt.person || '').trim();
  if (c && pe) return c + '-' + pe;
  return c || pe || String(idt.label || '').trim();
}

function saveIdentity(userId, idt) {
  var map = readIdentities();
  map[userId] = idt;
  var keys = Object.keys(map);
  if (keys.length > 300) keys.slice(0, keys.length - 300).forEach(function(k) { delete map[k]; });
  try { PropertiesService.getScriptProperties().setProperty(IDENTITY_KEY, JSON.stringify(map)); } catch (e) {}
  var label = identityLabel(idt);
  try { CacheService.getScriptCache().put('ULN_' + userId, JSON.stringify({ name: label, source: 'bind' }), 86400 * 7); } catch (e) {}
  return label;
}

function bindIdentity(userId, patch) {
  var cur = readIdentities()[userId] || {};
  var idt = { company: cur.company || '', person: cur.person || '' };
  if (patch.company) idt.company = patch.company;
  if (patch.person) idt.person = patch.person;
  return saveIdentity(userId, idt);
}

var IDT_BAD = /登記|新增|修改|改成|改為|送燈|安奉|軟體|電腦|總燈數|\d+\s*盞|\d\s*[*×xX]\s*\d|OLED|LED/i;
var IDT_RE = /(我的公司|我公司|公司名稱|公司|廠商|業務窗口|業務|經辦人|經辦|窗口|人員|姓名|名字叫|我叫|我是)(?:名稱)?\s*[:：]?\s*([^\s，,、；;:：]{1,20})/g;

// 解析「我公司 亞盛燈業 我叫 李小華」「我是 亞盛燈業 的 李小華」「我是聖文」等寫法
// 回傳 {company, person}；若這句看起來像登記內容（有殘字或含燈位關鍵字）則回傳 null
function parseIdentityCommand(text) {
  var t = String(text || '').trim();
  if (!t || t.length > 44) return null;

  var both = t.match(/^(?:我是|我叫)?\s*(.{1,20}?)\s*的\s*(.{1,20})$/);
  if (both && !IDT_BAD.test(both[1] + both[2]) && !/\s/.test(both[1] + both[2])) {
    return { company: both[1].trim(), person: both[2].trim() };
  }

  var res = {}, rest = t, hit = 0, re, m;
  IDT_RE.lastIndex = 0;
  var pieces = [];
  while ((m = IDT_RE.exec(t)) !== null) {
    var key = m[1], val = m[2].trim();
    if (IDT_BAD.test(val)) return null;
    if (PERSON_KEYS[key]) res.person = val; else res.company = val;
    pieces.push(m[0]);
    hit++;
  }
  if (!hit) return null;

  rest = t;
  pieces.forEach(function(seg) { rest = rest.replace(seg, ''); });
  rest = rest.replace(/[\s，,、；;:：]+/g, '');
  if (rest.length > 0) return null;   // 句子上還有別的內容 → 比較像是登記，不當身分綁定處理
  return res;
}

// 回傳該 LINE 使用者要帶入的業務欄文字（来源：綁定 > LINE 暱稱）
function getUserIdentity(userId) {
  var uid = String(userId || '');
  var bound = readIdentities()[uid];
  var label = identityLabel(bound);
  if (label) return { name: label, source: 'bind' };
  if (!uid || CONFIG.AGENT_FROM_LINE_PROFILE === false) return { name: '', source: 'none' };

  var ck = 'ULN_' + uid;
  try {
    var c = CacheService.getScriptCache().get(ck);
    if (c) { var j = JSON.parse(c); return { name: j.name || '', source: j.source || 'line' }; }
  } catch (e) {}
  var nick = fetchLineDisplayName(uid);
  try { CacheService.getScriptCache().put(ck, JSON.stringify({ name: nick, source: nick ? 'line' : 'none' }), 86400 * 3); } catch (e) {}
  return { name: nick, source: nick ? 'line' : 'none' };
}

function fetchLineDisplayName(userId) {
  try {
    var resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
      headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN, 'Content-Type': 'application/json' },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      console.log('🔇 取得 LINE 姓名失敗 HTTP ' + resp.getResponseCode() + '（未加好友／未同意授權／已封鎖）');
      return '';
    }
    return String(JSON.parse(resp.getContentText()).displayName || '').trim();
  } catch (e) {
    console.error('fetchLineDisplayName 異常: ' + e.toString());
    return '';
  }
}

// ==================== 2.4 回覆策略（只回「會動到資料」的訊息） ====================
// 腳本屬性 REPLY_MODE = ALL 可恢復「全部回覆」（含查詢與 AI 判定為 READ 的訊息）
function isReplyAllMode() {
  return /^(ALL|全部|所有|ON)$/i.test(String(CONFIG.REPLY_MODE || '').trim());
}

// 查詢類語氣（不含寫入動詞）：這種訊息不需要打扰使用者，也不需要花 Gemini 配額
function isReadOnlyAsking(text) {
  var t = String(text || '');
  var query = /查詢|查一下|查查看|幫我查|帮我查|有哪些|有幾盞|有幾種|幾盞|多少盞|總共|共計|列出|列一下|清單|統計|回報|查無|查一下/;
  var write = /新增|登記|補一筆|修改|改成|改為|換成|調整|送燈|安奉|強制新增|异动|異動/;
  return query.test(t) && !write.test(t);
}

// 與燈位資料無關的閒聊／雜訊（天氣、問候、語音轉文字的無關內容）
function looksLikeLightData(text) {
  var t = String(text || '');
  if (/盞|燈|廟|宮|壇|殿|祠|寺|庵|堂|光明|太歲|文昌|財神|月老|觀音|神尊|城隍|媽祖|聖母|福德|土地公|關帝|帝君|王爺|祖師|保生|玄天|立燈/.test(t)) return true;
  if (/登記|新增|修改|改成|改為|換成|調整|補一筆|安奉|業務|經辦|軟體|電腦/.test(t)) return true;
  if (/OLED|LED|研華|廟幫手|廟管家/i.test(t)) return true;
  if (/\d{1,2}\s*[*×xX＊]\s*\d{1,2}/.test(t)) return true;  // 5*7 / 4*5 這類規格寫法
  if (/[\u4e00-\u9fa5]{2,6}\s*[-－—]\s*[\u4e00-\u9fa5]{2,20}/.test(t) && /\d/.test(t)) return true;  // 「聖文-石岡子乾元宮 ... 2112」
  return false;
}

// 決定這則訊息要處理、回指令、還是完全沉默
// hasDraft：該使用者有 10 分鐘內的補件草稿；chatOn：處於聊天模式（免喚醒字、問句都回）
function decideFlow(text, hasDraft, chatOn) {
  var t = String(text || '').trim();
  if (!t) return { mode: 'silent', reason: '空訊息' };
  if (isHelpCommand(t)) return { mode: 'command', handler: 'help' };
  if (isModelCommand(t)) return { mode: 'command', handler: 'model' };

  var stripped = stripWakeWord(t);
  var woke = stripped !== null;
  var body = woke ? stripped : t;

  var cc = parseChatCommand(body) || (woke ? null : parseChatCommand(t));
  if (cc) return { mode: 'chatcmd', cmd: cc };

  if (!woke && !hasDraft && !chatOn && !isReplyAllMode()) {
    return { mode: 'silent', reason: '未帶喚醒字（' + (CONFIG.WAKE_WORDS || []).join('/') + '）' };
  }
  if (!body) return { mode: 'empty', text: '' };
  if (!chatOn && isReadOnlyAsking(body) && !isReplyAllMode()) {
    return { mode: 'silent', reason: '查詢類，請直接用瀏覽器開試算表' };
  }
  return { mode: 'ai', text: body, woke: woke, chat: !!chatOn };
}

// AI 判定後的最終把關：只有寫入類意圖（含寫入失敗的提示）才回訊
function shouldReplyFor(action) {
  var a = String(action || '').toUpperCase();
  if (isReplyAllMode()) return true;
  return a === 'CREATE' || a === 'UPDATE';
}

// ==================== 2.5 LINE 端模型管理指令 (/model, 模型) ====================
function isModelCommand(text) {
  var lower = text.toLowerCase().trim();
  return lower.indexOf('/model') === 0 || lower.indexOf('模型') === 0 || lower.indexOf('查模型') === 0;
}

function handleModelCommand(text) {
  var parts = text.split(/\s+/);
  var manualModel = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');
  var candidates = getCandidateModels();

  if (parts.length === 1 || /^(list|清單)$/i.test(parts[1])) {
    var reply = '🤖 【Gemini 模型設定狀態】\n━━━━━━━━━━━━━━\n';
    if (manualModel) {
      reply += '📌 當前模式：手動鎖定\n🎯 使用模型：' + manualModel + '（失效時自動退回清單其他模型）\n\n';
    } else {
      reply += '📌 當前模式：自動優先（依序尝试）\n🎯 目前優先：' + candidates[0] + '\n\n';
    }
    reply += '📋 偵測到此金鑰可用模型：\n';
    candidates.slice(0, 8).forEach(function(m, idx) {
      reply += (idx + 1) + '. ' + m + (m === (manualModel || CONFIG.GEMINI_MODEL) ? ' ← 使用中' : '') + '\n';
    });
    reply += '━━━━━━━━━━━━━━\n💡 切換：/model <模型名>（例：/model gemini-3.6-flash）\n💡 恢復預設：/model auto\n💡 完整使用說明：/help';
    return reply;
  }

  var target = parts[1].trim();
  if (/^(auto|reset|自動|預設)$/i.test(target)) {
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_MODEL');
    return '🔄 已恢復【自動模式】：將依「預設模型 > 最新可用模型清單」順序嘗試調用。';
  }

  PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL', target);
  return '✅ 已切換模型為：「' + target + '」\n下一則訊息立即生效。若此模型失效，系統會自動退回其他可用模型並繼續回覆。';
}

// ==================== 2.6 LINE 端使用說明指令 (/help, 使用方式, 說明) ====================
function isHelpCommand(text) {
  var lower = text.toLowerCase().trim();
  return lower === '/help' || lower === 'help' || lower === '/?' || lower === '?' || lower === '？' ||
         lower === '說明' || lower === '使用說明' || lower === '使用方式' || lower === '功能' ||
         lower === '菜單' || lower === '操作說明';
}

function handleHelpCommand() {
  var wake = (CONFIG.WAKE_WORDS && CONFIG.WAKE_WORDS[0]) || '小幫手';
  var req = (CONFIG.REQUIRED_FIELDS || []).map(function(k) { return CONFIG.FIELD_LABELS[k]; }).join('、');
  var lines = [
    '⛩️ 【光明燈管理系統｜使用方式】',
    '━━━━━━━━━━━━━━',
    '🔔 先喊「' + wake + '」我才會處理（也可用：' + (CONFIG.WAKE_WORDS || []).join('／') + '）',
    '',
    '📝 1. 新增登記（必填：' + req + '）',
    '　 例：' + wake + ' 聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前 軟體其他 電腦研華',
    '　 例：' + wake + ' 我要登記，媽祖廟新增財神燈500盞，業務王小明',
    '　 ⚠️ 必填欄位缺任何一項都不會上表，我會回覆「還需要什麼」格式',
    '　 👤 業務預設＝' + (CONFIG.DEFAULT_AGENT || '（未設定，會回問）') + '；沒特別寫就上這個名字',
    '　 👤 别家的人綁一次就蓋掉預設，之後不用每次寫業務：',
    '　　　 例：' + wake + ' 我是名典　／　' + wake + ' 我公司 亞盛燈業 我叫 李小華',
    '　　　 綁兩段時業務欄顯示「亞盛燈業-李小華」；換人或換公司再傳一次蓋掉',
    '　 👤 單筆想填別家：直接寫在句首，例：' + wake + ' 甫穎-仁武保安宮 ...',
    '　 ⏳ 收到提示後 ' + Math.round(CONFIG.DRAFT_TTL / 60) + ' 分鐘內直接補（例：' + wake + ' 規格5*7 OLED琥珀色），我會自動合併',
    '　 ⚠️ 只要廟名對到既有資料，我會把那些列列出來請你確認（不會靜悄悄地亂加或亂擋）',
    '　 ⚠️ 同廟＋同規格＋同承攬商 → 視為重複，請改用「修改」；確定是第二批請加「強制新增」',
    '　 ⚠️ 同廟同規格但承攬商不同 → 當成另一張單，直接新增；同名廟多會館 → 會要你寫全名',
    '',
    '✏️ 2. 修改（需含「廟宇」＋「規格」才能定位）',
    '　 例：' + wake + ' 修改 新化武廟 4*5 OLED琥珀色 總燈數變成5238盞 軟體冠緯 電腦研華',
    '　 例：' + wake + ' 金六結福德廟 5*7 OLED 送燈改成 國115/03/01前',
    '　 ✔️ 只講要改的欄位就好，沒提到的欄位不會被清空',
    '　 ✔️ 該廟只有一種規格時，可直接說「修改 新化武廟 總燈數變成5238盞」',
    '',
    '🔍 3. 查詢：建議直接用瀏覽器開試算表（LINE 端查詢不回訊）',
    '',
    '💬 4. 聊天模式（會開始每句都回你）',
    '　 ' + wake + ' 聊天　　　　進入（之後免打喚醒字）',
    '　 ' + wake + ' 結束聊天　　離開並清空對話記憶',
    '　 ' + wake + ' 角色　　　　看可選角色；' + wake + ' 角色 文案助手　換角色',
    '　 目前角色：' + (CONFIG.CHAT.DEFAULT_ROLE || '操作說明') + '（回答依操作手冊，上限 ' + chatCharLimit() + ' 字；' + Math.round(CONFIG.CHAT.MODE_TTL / 60) + ' 分鐘沒動自動關閉）',
    '',
    '🤖 5. AI 模型管理',
    '　 /model　　　　　查看目前模型與可用清單',
    '　 /model <模型名>　切換指定模型',
    '　 /model auto　　　恢復自動模式',
    '',
    '📋 可辨識欄位：業務｜廟宇｜規格｜總燈數｜送燈時間｜軟體｜電腦｜備註',
    '　 ✍️ 其他資訊都放備註欄：明確寫「備註：……」才會進 H 欄，例：…500盞 國10/17前 備註：分兩批送',
    '💡 送燈時間保留您的寫法（國10/17前、12/10or12/17、已送燈），系統不改寫 E 欄',
    '　　・「國」=國曆、「民國115」=西元2026；寫「農曆/舊曆」系統不敢換算',
    '　　・推得出的國曆日期會以〔國2026-10-17〕標記自動併進備註欄（關掉：腳本屬性 DATE_TAG_IN_REMARK=OFF）',
    '💡 合計／小計公式請放最右側欄（I 欄以後），機器人只寫 A~H，絕不動到你的公式',
    '',
    '🔕 回覆規則：只在「新增／修改」「資料不全需補件」或「聊天模式進行中」時回訊',
    '　 未帶喚醒字、查詢、閒聊一律不回覆（聊天模式除外）',
    '　 想恢復全部回覆：GAS 專案設定 → 腳本屬性 → 新增 REPLY_MODE = ALL',
    '━━━━━━━━━━━━━━',
    '✨ 帶「' + wake + '」開頭直接輸入登記內容；輸入 /help 隨時回來看說明。'
  ];
  return lines.join('\n');
}

// 向 Gemini 查詢這把金鑰實際可用的 flash 系列模型（結果快取 6 小時）
function getCandidateModels() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('GEMINI_MODELS');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  var fallback = [CONFIG.GEMINI_MODEL].concat(['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3-flash-preview']);
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + CONFIG.GEMINI_API_KEY + '&pageSize=200';
    var resp = UrlFetchApp.fetch(url, { "method": "get", "muteHttpExceptions": true });
    if (resp.getResponseCode() === 200) {
      var models = JSON.parse(resp.getContentText()).models || [];
      var usable = [];
      for (var i = 0; i < models.length; i++) {
        var name = (models[i].name || '').replace('models/', '');
        var methods = models[i].supportedGenerationMethods || [];
        if (methods.indexOf('generateContent') !== -1 && /flash/i.test(name) &&
            !/image|tts|audio|live|vision|embedding|robotics|computer|omni|lyria|native/i.test(name)) {
          usable.push(name);
        }
      }
      if (usable.length > 0) {
        var list = dedupe(usable.sort(modelVersionDesc));
        cache.put('GEMINI_MODELS', JSON.stringify(list), 21600);
        return list;
      }
    }
  } catch (e) {
    console.error('getCandidateModels 異常: ' + e.toString());
  }
  return dedupe(fallback);
}

function dedupe(arr) {
  var seen = {};
  return arr.filter(function(x) { if (!x || seen[x]) return false; seen[x] = true; return true; });
}

function modelVersionDesc(a, b) {
  function score(n) {
    var m = n.match(/(\d+(?:\.\d+)?)/);
    var v = m ? parseFloat(m[1]) : 0;
    return v + (/preview|exp|latest/i.test(n) ? -0.001 : 0);
  }
  return score(b) - score(a);
}

// ==================== 3. Gemini API 串接與語意分析 ====================
function analyzeMessageWithGemini(text) {
  
  var prompt = "你是一個宮廟「光明燈規格管理」的自動化助手。請分析使用者的輸入，並輸出為 JSON 格式（不要包含任何 ```json 字樣或 Markdown 外殼）。\n\n" +
               "【使用者輸入】: \"" + text + "\"\n\n" +
               "【輸出 JSON 欄位規範】:\n" +
               "1. action: 必須是 'CREATE'（新增）、'UPDATE'（修改）或 'READ'（查詢）其中之一。\n" +
               "2. details: 擷取核心資訊物件，包含以下欄位（若無提及則填 null，不可自行編造）:\n" +
               "   - agent: 業務 / 經辦人姓名（例：聖文、甫穎、冠緯、家森、冠宇）\n" +
               "   - temple: 廟宇 / 廟方名稱（例：石岡子乾元宮、桃園廣盛壇、新化武廟）\n" +
               "   - spec: 規格，請「原樣保留」尺寸、燈種與顏色（例：5*7 OLED琥珀色、4*5 OLED、7*9 OLED琥珀色）\n" +
               "   - total_count: 總燈數（必須是純數字，移除千分逗號；「2,112盞」請填 2112）\n" +
               "   - delivery_text: 送燈日期，請「原樣保留」使用者的寫法，包含「國」「前」「or」或「已送燈」等字樣（例：國10/17前、12/10or12/17、國115-02/09前、已送燈）。切勿自行換算成 YYYY-MM-DD。\n" +
               "   - software: 軟體（例：廟幫手、冠緯、家森、冠宇、其他、廟管家）\n" +
               "   - computer: 電腦／硬體（例：研華、廟、研華*3）\n" +
               "   - remark: 備註。只在使用者明確寫出「備註：…」或「另外說明：…」時才填，內容照原樣保留；\n" +
               "             沒有明示一律填 null，不可把整句登記內容或你自己推測的話寫進備註\n\n"
               "【重要規則】:\n" +
               "A. 使用者常把業務與廟名用「-」「－」「—」連寫（例：「聖文-石岡子乾元宮」），此時「-」之前是 agent，之後是 temple，必須拆開。\n" +
               "B. 「國115/02/09」屬民國年、「國10/17」屬國曆月日，兩者都只是日期寫法差異，原字串照抄進 delivery_text。\n" +
               "C. 同一間廟可能同時有 4*5 與 5*7 兩種規格，因此 spec 是重要的定位欄位，有提到就一定要填。\n" +
               "D. 語意為「查詢/有哪些/帮我查」時 action='READ'；「改成/調整/補登記」時 action='UPDATE'；其餘登記類描述為 'CREATE'。\n\n" +
               "【範例 1：新增（業務與廟名連寫）】:\n" +
               "輸入:「聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前 軟體其他 電腦研華 備註：分兩批送，第二批國12/01前」\n" +
               "輸出: {\"action\":\"CREATE\",\"details\":{\"agent\":\"聖文\",\"temple\":\"石岡子乾元宮\",\"spec\":\"5*7 OLED琥珀色\",\"total_count\":2112,\"delivery_text\":\"國10/17前\",\"software\":\"其他\",\"computer\":\"研華\",\"remark\":\"分兩批送，第二批國12/01前\"}}\n\n" +
               "【範例 2：新增（口語敘述）】:\n" +
               "輸入:「業務王小明登記，媽祖廟要新增財神燈500盞，預計10月15送燈」\n" +
               "輸出: {\"action\":\"CREATE\",\"details\":{\"agent\":\"王小明\",\"temple\":\"媽祖廟\",\"spec\":\"財神燈\",\"total_count\":500,\"delivery_text\":\"10月15\",\"software\":null,\"computer\":null,\"remark\":null}}\n\n" +
               "【範例 3：修改】:\n" +
               "輸入:「幫我修改新化武廟 4*5 OLED琥珀色 的總燈數變成5238盞，軟體冠緯、電腦研華」\n" +
               "輸出: {\"action\":\"UPDATE\",\"details\":{\"agent\":null,\"temple\":\"新化武廟\",\"spec\":\"4*5 OLED琥珀色\",\"total_count\":5238,\"delivery_text\":null,\"software\":\"冠緯\",\"computer\":\"研華\",\"remark\":null}}\n\n" +
               "【範例 4：查詢】:\n" +
               "輸入:「查一下金六結福德廟有哪些燈」\n" +
               "輸出: {\"action\":\"READ\",\"details\":{\"agent\":null,\"temple\":\"金六結福德廟\",\"spec\":null,\"total_count\":null,\"delivery_text\":null,\"software\":null,\"computer\":null,\"remark\":null}}";

  return askLLM({ user: prompt, wantJson: true, purpose: 'parse' });
}

// ==================== 3.5 LLM 傳輸層（GEMINI / OLLAMA） ====================
// purpose: 'parse'（登記解析）或 'chat'（聊天模式）；兩者可用不同供應者省額度
function providerFor(purpose) {
  var base = (CONFIG.LLM.PROVIDER || 'GEMINI').toUpperCase();
  var p = purpose === 'chat' ? (CONFIG.LLM.CHAT_PROVIDER || base) : CONFIG.LLM.PROVIDER;
  p = String(p).toUpperCase();
  return (p === 'OLLAMA' || p === 'GEMINI') ? p : 'GEMINI';
}

// opts = { system, user, history:[{role:'user'|'assistant',text}], wantJson, purpose, model }
function askLLM(opts) {
  var provider = providerFor(opts.purpose || 'parse');
  if (provider === 'OLLAMA') return askOllama(opts);
  return askGemini(opts);
}

function stripJsonFence(text) {
  return String(text || '').replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
}

function buildParsePromptText(opts) {
  var chunks = [];
  if (opts.system) chunks.push(opts.system);
  (opts.history || []).forEach(function (h) {
    chunks.push((h.role === 'assistant' ? '助理：' : '使用者：') + String(h.text));
  });
  chunks.push(opts.user || '');
  return chunks.join('\n');
}

function askGemini(opts) {
  var prompt = opts.system ? buildParsePromptText(opts) : opts.user;
  var payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  if (opts.wantJson) payload.generationConfig = { "responseMimeType": "application/json" };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  // 嘗試順序：指定模型 > 手動鎖定 > 預設模型 > 自動探測的最新可用模型
  var manualModel = CONFIG.GEMINI_MODEL;
  var order = dedupe([opts.model, manualModel, CONFIG.GEMINI_MODEL].concat(getCandidateModels())).slice(0, 6);
  var lastErr = '', lastCode = 0;

  for (var mi = 0; mi < order.length; mi++) {
    var tryUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + order[mi] + ":generateContent?key=" + CONFIG.GEMINI_API_KEY;
    var response = UrlFetchApp.fetch(tryUrl, options);
    var respCode = response.getResponseCode();
    var bodyText = response.getContentText();

    if (respCode !== 200) {
      lastErr = '模型 ' + order[mi] + ' (HTTP ' + respCode + '): ' + bodyText.substring(0, 200);
      lastCode = respCode;
      // 429/503 屬暫時性（配額或壅塞），每個模型各自獨立，繼續換下一個並稍作等待
      if ((respCode === 429 || respCode === 500 || respCode === 503 || respCode === 504) && mi < order.length - 1) {
        Utilities.sleep(1200);
      }
      continue;
    }
    try {
      var jsonResponse = JSON.parse(bodyText);
      if (!jsonResponse.candidates || !jsonResponse.candidates[0]) {
        lastErr = '模型 ' + order[mi] + ' 無候選回覆: ' + bodyText.substring(0, 200);
        continue;
      }
      var parts = jsonResponse.candidates[0].content.parts || [];
      var aiText = parts.map(function (p) { return p.text || ''; }).join('').trim();
      return opts.wantJson ? JSON.parse(stripJsonFence(aiText)) : aiText;
    } catch (e) {
      lastErr = '模型 ' + order[mi] + ' 解析失敗: ' + e.toString();
      continue;
    }
  }
  if (lastCode === 503 || lastCode === 500 || lastCode === 504) {
    throw new Error('AI 模型暫時壅塞（已嘗試 ' + order.length + ' 個模型），請稍候 1 分鐘重傳；或輸入「/model gemini-3.1-flash-lite」改用較穩定模型。');
  }
  if (lastCode === 429) {
    throw new Error('AI 免費配額暫時用盡（已嘗試 ' + order.length + ' 個模型），請稍後重試或輸入「/model gemini-3.1-flash-lite」。');
  }
  throw new Error('Gemini API 異常: ' + lastErr);
}

// Ollama：https://ollama.com/api/chat，Bearer token；免費層 1 併發
function askOllama(opts) {
  var key = CONFIG.OLLAMA_API_KEY || '';
  if (!key || key === 'YOUR_OLLAMA_API_KEY') {
    throw new Error('要用 Ollama 請在 GAS「專案設定 → 腳本屬性」新增 OLLAMA_API_KEY（到 ollama.com/settings/keys 產生），不要寫進程式檔');
  }

  var base = String(CONFIG.LLM.OLLAMA_BASE_URL).replace(/\/+$/, '');
  var msgs = [];
  if (opts.system) {
    msgs.push({ role: 'system', content: opts.system + (opts.wantJson ? '\n【輸出硬性要求】只輸出 JSON 物件本身，不要 markdown 程式碼塊、不要任何說明文字。' : '') });
  }
  (opts.history || []).forEach(function (h) {
    msgs.push({ role: h.role === 'assistant' ? 'assistant' : 'user', content: String(h.text) });
  });
  msgs.push({ role: 'user', content: opts.wantJson && !opts.system ? buildParsePromptText(opts) : (opts.user || '') });

  var payload = {
    model: opts.model || CONFIG.LLM.OLLAMA_MODEL,
    messages: msgs,
    stream: false,
    options: { temperature: opts.wantJson ? 0 : 0.6 }
  };
  if (opts.wantJson) payload.format = 'json';

  var resp = UrlFetchApp.fetch(base + '/api/chat', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode(), text = resp.getContentText();
  if (code !== 200) {
    throw new Error('Ollama API HTTP ' + code + '：' + String(text).substring(0, 160) + '｜' + CONFIG.LLM.OLLAMA_TIMEOUT_HINT);
  }
  var json;
  try { json = JSON.parse(text); } catch (e) {
    throw new Error('Ollama 回覆解析失敗：' + String(text).substring(0, 160));
  }
  var content = (json && json.message && json.message.content) ? String(json.message.content) : '';
  if (!opts.wantJson) return content.trim();
  try {
    return JSON.parse(stripJsonFence(content));
  } catch (e2) {
    throw new Error('Ollama 沒回出可解析的 JSON（模型太輕或不聽 format 參數）。建議把解析改回 GEMINI、只用 Ollama 聊天：腳本屬性 LLM_PROVIDER=GEMINI、CHAT_PROVIDER=OLLAMA');
  }
}

// 三家连通性自检：在 GAS 编辑器执行，结果看「执行记录」
function testLLMProviders() {
  var rows = [];
  ['GEMINI', 'OLLAMA'].forEach(function (provider) {
    var t0 = new Date().getTime(), out;
    try {
      var r = askLLM({ system: '你是连线测试。只回覆两个字：OK', user: 'ping', wantJson: false, provider: provider });
      out = '✅ ' + String(r).substring(0, 20);
    } catch (e) {
      out = '❌ ' + e.message;
    }
    rows.push(provider + '（' + ((new Date().getTime()) - t0) + 'ms）：' + out);
  });
  Logger.log(rows.join('\n'));
  return rows.join('\n');
}


// ==================== 3.6 聊天模式：開關、角色、操作手冊、對話記憶 ====================
var CHAT_STATE_KEY = 'USER_CHAT';

function readChatStates() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(CHAT_STATE_KEY) || '{}'); } catch (e) { return {}; }
}

function writeChatStates(map) {
  var keys = Object.keys(map);
  if (keys.length > 300) {
    keys.sort(function (a, b) { return (map[a].at || 0) - (map[b].at || 0); });
    keys.slice(0, keys.length - 300).forEach(function (k) { delete map[k]; });
  }
  try { PropertiesService.getScriptProperties().setProperty(CHAT_STATE_KEY, JSON.stringify(map)); } catch (e) {}
}

// 聊天模式超過 MODE_TTL 沒有動作就自動關（不會有人忘了關，之後每句話都被回覆）
function getChatState(userId) {
  var st = readChatStates()[String(userId)];
  if (!st || !st.on) return null;
  if ((new Date().getTime()) - (st.at || 0) > CONFIG.CHAT.MODE_TTL * 1000) return null;
  return st;
}

function setChatState(userId, on, role) {
  var map = readChatStates(), k = String(userId), cur = map[k] || {};
  map[k] = on
    ? { on: true, role: role || cur.role || CONFIG.CHAT.DEFAULT_ROLE, at: new Date().getTime() }
    : { on: false, role: cur.role || CONFIG.CHAT.DEFAULT_ROLE, at: 0 };
  writeChatStates(map);
  return map[k];
}

// ---- 對話記憶（CacheService 上限 600 秒，與 HISTORY_TTL 搭配）----
function historyKey(userId) { return 'CHATHIST_' + userId; }

function loadHistory(userId) {
  try {
    var c = CacheService.getScriptCache().get(historyKey(userId));
    if (c) return JSON.parse(c) || [];
  } catch (e) {}
  return [];
}

function pushHistory(userId, userText, aiText) {
  var h = loadHistory(userId);
  h.push({ role: 'user', text: String(userText).substring(0, 500) });
  h.push({ role: 'assistant', text: String(aiText).substring(0, 1200) });
  var max = Math.max(2, CONFIG.CHAT.HISTORY_TURNS * 2);
  if (h.length > max) h = h.slice(h.length - max);
  try { CacheService.getScriptCache().put(historyKey(userId), JSON.stringify(h), CONFIG.CHAT.HISTORY_TTL); } catch (e) {}
  return h;
}

function clearHistory(userId) {
  try { CacheService.getScriptCache().remove(historyKey(userId)); } catch (e) {}
}

// ---- 操作手冊：Google 文件 → 「操作手冊」分頁 → /help 內建文案 ----
function manualItems() {
  var cacheKey = 'CHATMANUAL';
  try {
    var c = CacheService.getScriptCache().get(cacheKey);
    if (c) return JSON.parse(c);
  } catch (e) {}

  var items = manualFromDoc();
  if (!items.length) items = manualFromSheet();
  if (!items.length) items = [{ topic: '系統內建說明', body: handleHelpCommand(), sample: '' }];

  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(items).substring(0, 95000), CONFIG.CHAT.HISTORY_TTL); } catch (e) {}
  return items;
}

function manualFromDoc() {
  var id = CONFIG.CHAT.MANUAL_DOC_ID;
  if (!id) return [];
  try {
    var text = String(DocumentApp.openById(id).getBody().getText());
    var paras = text.split(/\n\s*\n/), out = [];
    paras.forEach(function (para) {
      var lines = String(para).split('\n').map(function (l) { return l.trim(); }).filter(function (l) { return l; });
      if (!lines.length) return;
      if (/合計|小計|總計/.test(lines[0])) return;
      out.push({ topic: lines[0].substring(0, 40), body: lines.join('\n'), sample: '' });
    });
    return out;
  } catch (e) {
    console.log('讀 Google 文件手冊失敗（通常是沒啟用 DocumentApp 服務或未授權）：' + e.toString());
    return [];
  }
}

function manualFromSheet() {
  try {
    var sheet = getSpreadsheet().getSheetByName(CONFIG.CHAT.MANUAL_SHEET);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues(), out = [];
    for (var i = 1; i < data.length; i++) {
      var topic = String(data[i][0] || '').trim(), body = String(data[i][1] || '').trim(), sample = String(data[i][2] || '').trim();
      if (!topic && !body) continue;
      if (/合計|小計|總計/.test(topic)) continue;
      out.push({ topic: topic, body: body, sample: sample });
    }
    return out;
  } catch (e) {
    console.log('讀操作手冊分頁失敗：' + e.toString());
    return [];
  }
}

// 只挑與問句最相關的幾條進 prompt，省 token 也比較不會答非所問
function pickManual(question) {
  var items = manualItems();
  var q = String(question || '');
  if (items.length <= 4) return renderManual(items);

  var scored = items.map(function (it, idx) {
    var hay = it.topic + ' ' + it.body, score = 0;
    for (var i = 0; i + 2 <= hay.length; i++) {
      var gram = hay.substr(i, 2);
      if (q.indexOf(gram) !== -1) score += 1;
    }
    Object.keys(CONFIG.FIELD_LABELS).forEach(function (k) {
      if (q.indexOf(CONFIG.FIELD_LABELS[k]) !== -1 && (it.topic + it.body).indexOf(CONFIG.FIELD_LABELS[k]) !== -1) score += 4;
    });
    return { it: it, score: score, idx: idx };
  });
  scored.sort(function (a, b) { return b.score - a.score || a.idx - b.idx; });
  var hits = scored.filter(function (x) { return x.score > 0; }).slice(0, 5).map(function (x) { return x.it; });
  if (!hits.length) hits = scored.slice(0, 3).map(function (x) { return x.it; });   // 一個都沒命中，才給開頭幾條當背景
  return renderManual(hits);
}

function renderManual(items) {
  var out = items.map(function (it) {
    return '【' + (it.topic || '未訂標題') + '】' + it.body + (it.sample ? '\n　可複製範例：' + it.sample : '');
  }).join('\n');
  return out.length > CONFIG.CHAT.MANUAL_MAX_CHARS ? out.substring(0, CONFIG.CHAT.MANUAL_MAX_CHARS) + '…（手冊過長已截斷）' : out;
}

function chatSystemPrompt(role, question) {
  return (CONFIG.CHAT.ROLES[role] || CONFIG.CHAT.ROLES[CONFIG.CHAT.DEFAULT_ROLE]) + '\n\n' +
    '【這套系統的硬性規則】\n' +
    '1. 經辦人要傳「' + (CONFIG.WAKE_WORDS[0] || '小幫手') + '」開頭的訊息系統才會處理（聊天模式進行中則免）\n' +
    '2. 必填欄位：' + (CONFIG.REQUIRED_FIELDS || []).map(function (k) { return CONFIG.FIELD_LABELS[k]; }).join('、') + '；缺件不上表，系統會回問\n' +
    '3. 寫入的分頁是「' + CONFIG.SHEET_NAMES.LIGHT_MGMT + '」，欄位依序：' + CONFIG.HEADERS.join('、') + '；合計/小計公式請放 I 欄以後\n' +
    '4. 業務欄沒寫時：先用本人綁定的公司／人員，再退回預設「' + (CONFIG.DEFAULT_AGENT || '（未設定，會回問）') + '」\n' +
    '5. 系統只在「新增／修改／補件」時回訊；查詢建議直接用瀏覽器開試算表\n\n' +
    '【回覆長度硬性規定】全繁體中文，' + chatCharLimit() + '字以內（含標點）；' +
    '先講結論再條列重點，最多 4 條，不要寒暄、不要重複題目、不要列參考來源；' +
    '需要更長的說明就請對方再問\n\n' +
    '【操作手冊（只能以此為準；沒寫到的就說不確定，不要編造）】\n' + pickManual(question);
}

// 聊天回覆字數上限：腳本屬性 CHAT_MAX_CHARS 優先，抓 200~2000 中間值避免被人設成 0 或超長
function chatCharLimit() {
  var n = parseInt(CONFIG.CHAT.MAX_REPLY_CHARS, 10);
  if (isNaN(n)) n = 600;
  return Math.min(2000, Math.max(80, n));   // 夾住，避免被設成 0 或超長
}

// 真正回一句聊天；provider 由 CHAT_PROVIDER 決定（可與登記解析不同）
function chatAnswer(userId, question) {
  var st = getChatState(userId) || setChatState(userId, true);
  var role = st.role || CONFIG.CHAT.DEFAULT_ROLE;
  var system = chatSystemPrompt(role, question);

  var text = askLLM({ system: system, user: String(question).substring(0, 1500), history: loadHistory(userId), wantJson: false, purpose: 'chat' });
  text = clipChatReply(text);
  pushHistory(userId, question, text);
  setChatState(userId, true, role);   // 續命，避免講到一半自動關
  return '💬【' + role + '】\n' + text;
}

// 模型常常不聽字數指示，回傳前自己再硬截一道（保留行尾，標明被截斷）
function clipChatReply(text) {
  var limit = chatCharLimit();
  var t = String(text || '').replace(/\r/g, '').trim();
  if (!t) return '（沒什麼回應，再問一次看看）';
  if (t.length <= limit) return t;
  var cut = t.lastIndexOf('\n', limit);
  if (cut < Math.floor(limit * 0.6)) cut = limit;
  return t.substring(0, cut).trim() + '\n…（已限制 ' + limit + ' 字，要更詳細請再問一句）';
}

// ---- 聊天相關指令解析 ----
// 回傳 null 代表不是聊天指令；否則 {action:'on'|'off'|'roles'|'role', role}
function parseChatCommand(text) {
  var raw = String(text || '').trim();
  var squashed = raw.replace(/[\s　]+/g, '').toLowerCase();
  if (/^(\/chat|\/聊天|聊天|聊天模式|進入聊天|開聊天|開始聊天)$/.test(squashed)) return { action: 'on' };
  if (/^(\/chatoff|\/chat:off|\/結束聊天|結束聊天|離開聊天|退出聊天|結束對話|不聊了)$/.test(squashed)) return { action: 'off' };
  var m = raw.match(/^\/?\s*(?:角色|role)(?:\s+(.{1,20}))?$/i);
  if (m) return { action: m[1] ? 'role' : 'roles', role: String(m[1] || '').trim() };
  return null;
}

function handleChatCommand(userId, cmd) {
  var roleList = Object.keys(CONFIG.CHAT.ROLES);
  if (cmd.action === 'on') {
    var st = setChatState(userId, true);
    return '💬 已進入【聊天模式】（目前角色：' + st.role + '）\n' +
           '・往後每則訊息都會回你，不用再打「' + (CONFIG.WAKE_WORDS[0] || '小幫手') + '」\n' +
           '・登記類訊息仍會優先入表（例：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 石岡子乾元宮 5*7 OLED 2112盞）\n' +
           '・換角色：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 角色 清單　或　' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 角色 ' + (Object.keys(CONFIG.CHAT.ROLES)[1] || '') + '\n' +
           '・結束：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 結束聊天（或 ' + Math.round(CONFIG.CHAT.MODE_TTL / 60) + ' 分鐘沒動會自動關）';
  }
  if (cmd.action === 'off') {
    setChatState(userId, false);
    clearHistory(userId);
    return '🔕 已結束聊天模式，恢復只回「新增／修改／補件」\n輸入 /help 可再看使用方式。';
  }
  if (cmd.action === 'roles') {
    var cur = (getChatState(userId) || {}).role || CONFIG.CHAT.DEFAULT_ROLE;
    return '🎭 可選角色（目前：' + cur + '）\n' +
      roleList.map(function (r, i) {
        return '  ' + (i + 1) + '. ' + r + (r === cur ? ' ← 使用中' : '') + '：' + String(CONFIG.CHAT.ROLES[r]).substring(0, 28) + '…';
      }).join('\n') +
      '\n切換：' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' 角色 ' + roleList[1];
  }
  var want = cmd.role;
  var hit = null;
  roleList.forEach(function (r) { if (!hit && (r === want || r.indexOf(want) !== -1 || want.indexOf(r) !== -1)) hit = r; });
  if (!hit) return '❌ 沒有這個角色：「' + want + '」\n可選：' + roleList.join('／');
  setChatState(userId, true, hit);
  return '🎭 已切換角色：' + hit + '（聊天模式同時開啟中；結束請傳「結束聊天」）';
}

// ==================== 4. 核心業務邏輯：新增、修改、查詢 ====================
function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID') {
    throw new Error('尚未設定 SPREADSHEET_ID（GAS 專案設定 → 腳本屬性）');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

// createIfMissing=true 時會自動建立分頁（供 setupLightSheet 使用）
function getLightSheet(createIfMissing) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LIGHT_MGMT);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.LIGHT_MGMT);
    Logger.log('🆕 已建立分頁「' + CONFIG.SHEET_NAMES.LIGHT_MGMT + '」');
  }
  return sheet;
}

// 系統會用到的最右欄索引（新增欄位只改 CONFIG.COLUMNS 就好，不必各處同步）
function maxColumnIndex() {
  var max = 0;
  Object.keys(CONFIG.COLUMNS).forEach(function(k) { max = Math.max(max, CONFIG.COLUMNS[k]); });
  return max;
}

function sheetTimeZone() {
  try { return Session.getScriptTimeZone(); } catch (e) { return 'Asia/Taipei'; }
}

// 資料區（含表頭）；一律裁到系統管理的欄寬（目前 A~H），右側由使用者自放的合計/公式欄不受影響
function readData(sheet) {
  var width = maxColumnIndex();
  var values = sheet.getDataRange().getValues();
  var trimmed = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i].slice(0, width);
    for (var j = row.length; j < width; j++) row.push('');
    trimmed.push(row);
  }
  return trimmed;
}

// 系統管理的欄寬（A~H）若超過現有欄數，先自動補欄避免 getRange 越界
function ensureColumnWidth(sheet) {
  var need = maxColumnIndex();
  if (sheet.getMaxColumns() < need) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  }
}

// 於 GAS 編輯器手動執行一次：補齊欄位並寫入表頭
// 【上線執行這一支就好】分頁不存在就自動建立，並寫入 A~H 表頭、補齊欄位
// 重複執行安全：已有資料不會被清掉，只把第 1 列表頭蓋成同一組名稱
function setupLightSheet() {
  var sheet = getLightSheet(true);
  decorateSheet(sheet);
  var msg = '✅ 主表「' + CONFIG.SHEET_NAMES.LIGHT_MGMT + '」就緒（唯一寫入端，欄位 A~' +
            String.fromCharCode(64 + maxColumnIndex()) + '）：' + CONFIG.HEADERS.join(' | ');
  Logger.log(msg);
  return msg + '\n' + syncCompanyTabs();
}

function decorateSheet(sheet) {
  ensureColumnWidth(sheet);
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  sheet.setFrozenRows(1);
  try { sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold'); } catch (e) {}
}

// 各承攬商分頁（不含主表）：只讀投影，由主表 QUERY 產生
function companyTabNames() {
  var master = CONFIG.SHEET_NAMES.LIGHT_MGMT, out = [];
  Object.keys(CONFIG.SHEET_NAMES).forEach(function(k) {
    if (k === 'LIGHT_MGMT') return;
    var name = String(CONFIG.SHEET_NAMES[k] || '').trim();
    if (name && name !== master) out.push(name);
  });
  return out;
}

// 建立／更新各公司分頁：表頭 + A2 起自動投影主表資料，並設為「編輯時警告」避免被人手改
function syncCompanyTabs() {
  var names = companyTabNames();
  if (!names.length) return '（未設定公司分頁）';
  var master = CONFIG.SHEET_NAMES.LIGHT_MGMT, ss = getSpreadsheet(), made = [];
  var lastCol = String.fromCharCode(64 + maxColumnIndex());          // 目前 9 → I（含備註欄）

  names.forEach(function(name) {
    var tab = ss.getSheetByName(name) || ss.insertSheet(name);
    var width = maxColumnIndex();
    if (tab.getMaxRows() < 30) tab.insertRowsAfter(tab.getMaxRows(), 30 - tab.getMaxRows());
    tab.getRange(2, 1, Math.max(tab.getMaxRows() - 1, 1), width).clearContent();  // 只清投影區，不動表頭
    tab.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    try { tab.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold'); } catch (e) {}
    var q = 'where upper(A) contains upper("' + name + '")';
    tab.getRange(2, 1).setFormula("=QUERY('" + master + "'!A2:" + lastCol + ', "' + q + '", 0)');
    tab.setFrozenRows(1);
    try {
      var prot = tab.protect().setDescription('由主表自動投影，請勿手動編輯');
      prot.setWarningOnly(true);
    } catch (e) { console.log('設定保護範圍失敗（不影響投影）：' + e.toString()); }
    made.push(name);
  });

  var msg = '🪞 公司分頁已同步：' + made.join('、') + '（只讀投影，資料仍以「' + master + '」為準）';
  Logger.log(msg);
  return msg;
}

// 只寫表頭、不建立分頁（分頁已存在時用）
function setupLightHeader() {
  var sheet = getLightSheet();
  if (!sheet) throw new Error('找不到分頁「' + CONFIG.SHEET_NAMES.LIGHT_MGMT + '」，請改執行 setupLightSheet()');
  ensureColumnWidth(sheet);
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  Logger.log('✅ 表頭已寫入：' + CONFIG.HEADERS.join(' | '));
  return CONFIG.HEADERS;
}

// 依欄位索引取儲存格（統一回傳字串處理安全）
function cell(row, colEnum) {
  var v = row[colEnum - 1];
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    try { return Utilities.formatDate(v, sheetTimeZone(), 'yyyy-MM-dd'); } catch (e) { return String(v); }
  }
  return String(v).trim();
}

// 「合計燈數」等小計／標題列一律跳過：只掃 業務／廟宇／規格 三欄文字，
// 右側自放的合計公式欄與使用者自行處理的小計都不會被誤判
function isSummaryRow(row) {
  var C = CONFIG.COLUMNS;
  var text = [cell(row, C.AGENT), cell(row, C.TEMPLE), cell(row, C.SPEC)].join(' ');
  return /合計|總計|小計/.test(text);
}

function toNumber(v) {
  if (typeof v === 'number') return v;
  var n = parseInt(String(v === null || v === undefined ? '' : v).replace(/[^\d]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

function withComma(v) {
  var n = toNumber(v);
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// 規格正規化：5*7 / 5×7 / 5 7 / 大寫小寫 皆視為相同
function normSpec(v) {
  return String(v === null || v === undefined ? '' : v).toLowerCase().replace(/[\s　*×x✕＊＊]/g, '');
}

function normName(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\s　（）()、,，]/g, '');
}

function matchTemple(cellVal, keyword) {
  var a = normName(cellVal), b = normName(keyword);
  if (!a || !b) return false;
  return a === b || a.indexOf(b) !== -1 || b.indexOf(a) !== -1;
}

function matchSpec(cellVal, keyword) {
  var a = normSpec(cellVal), b = normSpec(keyword);
  if (!a || !b) return false;
  return a === b || a.indexOf(b) !== -1 || b.indexOf(a) !== -1;
}

function containsLoose(cellVal, keyword) {
  var a = normName(cellVal), b = normName(keyword);
  return !!a && !!b && (a.indexOf(b) !== -1 || b.indexOf(a) !== -1);
}

// 送燈日期正規化（僅作為右側輔助欄）：支援「國10/17前」「國115-02/09前」「12/10or12/17」「10月15日」「已送燈」
function parseDateKey(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (raw instanceof Date) {
    try { return Utilities.formatDate(raw, sheetTimeZone(), 'yyyy-MM-dd'); } catch (e) { return ''; }
  }
  var s = String(raw)
    .replace(/[\s　]/g, '')
    .replace(/[／]/g, '/')
    .replace(/[－—–]/g, '-')
    .replace(/日/g, '');
  s = s.replace(/月/g, '/');
  if (/已送燈|已送|已完|送完|已完成/.test(s)) return '';
  // 農曆（舊曆/陰曆/国曆以外的曆別）不做換算：GAS 沒有農曆對照表，硬推會推出錯的日期
  if (/農曆|农历|舊曆|旧历|阴历|陰曆|國曆初一|初[一二三四五六七八九十]{1,3}／|^農/.test(s)) return '';

  var y = 0, m = 0, d = 0;
  var ROC_BASE = 1911;  // 民國 115 = 西元 2026、116 = 2027
  var full = s.match(/(?:民國|民|國|西元|公元)?(\d{2,4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (full) {
    y = parseInt(full[1], 10); m = parseInt(full[2], 10); d = parseInt(full[3], 10);
    if (y < 1900) y += ROC_BASE;  // 115 -> 2026（4 位數一律視為西元）
  } else {
    var md = s.match(/(\d{1,2})\/(\d{1,2})/);
    if (md) {
      y = new Date().getFullYear(); m = parseInt(md[1], 10); d = parseInt(md[2], 10);
    }
  }
  if (!y || !m || !d || m < 1 || m > 12 || d < 1 || d > 31) return '';
  var dt = new Date(y, m - 1, d);
  if (isNaN(dt.getTime())) return '';
  try { return Utilities.formatDate(dt, sheetTimeZone(), 'yyyy-MM-dd'); } catch (e) { return y + '-' + m + '-' + d; }
}

var DATE_TAG_RE = /[〔\[]\s*國\s*[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}\s*[〕\]]\s*[，,、\s]*/g;

// 拿掉備註裡系統自動加的〔國…〕標記，避免重複堆疊
function stripDateTag(text) {
  return String(text === null || text === undefined ? '' : text).replace(DATE_TAG_RE, '').replace(/^[，,、\s]+/, '').trim();
}

// 備註欄 = 使用者的說明 + 系統推的國曆標記（放在最前面方便篩選）
function buildRemark(userRemark, deliveryText) {
  var note = stripDateTag(userRemark);
  if (!CONFIG.REMARK_DATE_TAG) return note;
  var iso = parseDateKey(deliveryText);
  if (!iso) return note;
  return note.indexOf(iso) !== -1 ? note : '〔國' + iso + '〕' + (note ? (note.charAt(0) === '，' || note.charAt(0) === '、' ? '' : ' ') + note : '');
}

// 依 details 組出一列寫入值（依照 CONFIG.COLUMNS 索引擺放，可補洞）
function buildRowValues(d) {
  var C = CONFIG.COLUMNS;
  var vals = [];
  vals[C.AGENT - 1] = d.agent || '';
  vals[C.TEMPLE - 1] = d.temple || '';
  vals[C.SPEC - 1] = d.spec || '';
  vals[C.TOTAL - 1] = toNumber(d.total_count);
  vals[C.DELIVERY - 1] = d.delivery_text || '';
  vals[C.REMARK - 1] = buildRemark(d.remark, d.delivery_text);
  vals[C.SOFTWARE - 1] = d.software || '';
  vals[C.COMPUTER - 1] = d.computer || '';

  for (var i = 0; i < vals.length; i++) {
    if (vals[i] === undefined) vals[i] = '';
  }
  return vals;
}

function formatRecordLine(row, prefix) {
  var C = CONFIG.COLUMNS;
  var head = (cell(row, C.AGENT) || '未註明') + '｜' + (cell(row, C.TEMPLE) || '未註明');
  return (prefix || '') + '🏮 ' + head + '\n' +
         '    規格：' + (cell(row, C.SPEC) || '未註明') + '｜燈數：' + withComma(cell(row, C.TOTAL)) + ' 盞\n' +
         '    送燈：' + (cell(row, C.DELIVERY) || '未註明') +
         '｜軟體：' + (cell(row, C.SOFTWARE) || '未註明') +
         '｜電腦：' + (cell(row, C.COMPUTER) || '未註明') +
         (cell(row, C.REMARK) ? '\n    備註：' + cell(row, C.REMARK) : '');
}

function handleDataRouting(aiResult, originalText, userId) {
  var sheet = getLightSheet();
  if (!sheet) return "【⚠️ 系統錯誤】找不到分頁「" + CONFIG.SHEET_NAMES.LIGHT_MGMT + "」。\n處置：到 GAS 編輯器把函式選單切到 setupLightSheet → 執行一次（會自動建立分頁並寫好 A~H 表頭）。";
  if (!aiResult || !aiResult.details) return "【⚠️ 無法解析】AI 未回傳有效欄位，請換一種說法重試（輸入 /help 看範例）。";

  ensureColumnWidth(sheet);
  var action = String(aiResult.action || '').toUpperCase();
  var d = aiResult.details;

  if (action === 'CREATE') return doCreate(sheet, d, originalText, userId);
  if (action === 'UPDATE') return doUpdate(sheet, d, originalText);
  if (action === 'READ') return doRead(sheet, d, originalText);
  return "【⚠️ 無法辨識意圖】（" + (aiResult.action || '無') + "）\n💡 輸入 /help 查看使用方式";
}

// ------------------ 1. 新增 ------------------
// 只要有碰到同廟，一律把該廟現有的列列出來給對方確認，避免「同名不同館」被亂擋或亂加
function doCreate(sheet, d, originalText, userId) {
  var C = CONFIG.COLUMNS;
  if (!d.temple && !d.spec) {
    return "【⚠️ 新增失敗】至少要寫「廟宇」與「規格」，例：\n聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前 軟體其他 電腦研華";
  }

  var data = readData(sheet);
  var sameTemple = [];   // 廟名（模糊）命中的列號
  for (var i = 1; i < data.length; i++) {
    if (isSummaryRow(data[i])) continue;
    if (d.temple && matchTemple(cell(data[i], C.TEMPLE), d.temple)) sameTemple.push(i + 1);
  }
  var force = /強制新增/.test(originalText || '');

  // 情況一：這廟从没登記過 → 直接新增
  if (!sameTemple.length) return appendNewRow(sheet, d, '🆕 新廟名（表上是第一筆）');

  // 情況二：同廟但規格完全不同 → 視為另一種規格，照常新增並列出既有規格讓你知道
  var exact = [];
  sameTemple.forEach(function (r) { if (matchSpec(cell(data[r - 1], C.SPEC), d.spec)) exact.push(r); });
  if (!force && !exact.length) {
    return appendNewRow(sheet, d, 'ℹ️ 該廟既有規格：' + sameTemple.map(function (r) {
      return (cell(data[r - 1], C.SPEC) || '未填') + '（' + withComma(cell(data[r - 1], C.TOTAL)) + ' 盞）';
    }).join('、') + '｜本次是新規格，已另外加一列');
  }

  // 情況三：同廟同規格命中兩列以上 → 不能猜，列清單請對方指明（同时存草稿）
  if (!force && exact.length > 1) {
    var optLines = exact.map(function (r) {
      return '    · 第 ' + r + ' 列：' + (cell(data[r - 1], C.TEMPLE) || '未填') +
             '｜' + (cell(data[r - 1], C.SPEC) || '未填規格') +
             '｜' + (cell(data[r - 1], C.AGENT) || '未填業務') +
             '｜' + withComma(cell(data[r - 1], C.TOTAL)) + ' 盞';
    }).join('\n');
    if (userId) saveDraft(userId, { raw: originalText || '', details: d });
    return '【⚠️ 同名廟有 ' + exact.length + ' 列相同規格，先不動資料】\n' + optLines + '\n' +
           '➡️ 請把廟名寫完整再傳一次（含分館／地址），例：\n' +
           '    ' + (CONFIG.WAKE_WORDS[0] || '小幫手') + ' ' +
           (cell(data[exact[0] - 1], C.AGENT) || '聖文') + '-' + (cell(data[exact[0] - 1], C.TEMPLE) || d.temple) +
           ' ' + (d.spec || (cell(data[exact[0] - 1], C.SPEC) || '5*7 OLED')) + ' ' +
           (d.total_count || '200') + '盞\n' +
           '💡 這確實是要另起的一筆（例如第二批），訊息尾端加「強制新增」即可直接入表';
  }

  // 情況四：正好命中一列
  var existRow = data[exact[0] - 1];
  var existAgent = cell(existRow, C.AGENT), newAgent = String(d.agent || '').trim();
  var sameAgent = !newAgent || !existAgent || containsLoose(existAgent, newAgent);

  if (!force && sameAgent) {
    return '【⚠️ 疑似重複登記】第 ' + exact[0] + ' 列已有同廟同規格紀錄：\n' +
           formatRecordLine(existRow, '    ') + '\n' +
           '✏️ 若是要改數量，請改說：修改 ' + (cell(existRow, C.TEMPLE) || d.temple) + ' ' +
           (cell(existRow, C.SPEC) || d.spec || '') + ' 總燈數變成 2500 盞\n' +
           '💡 若這真的是第二批，請在訊息尾端加上「強制新增」再送一次。';
  }

  // 同廟同規格但承攬商不同 → 當成另一張單新增（同一間廟兩家共填時不會互相擋）
  if (!force) {
    return appendNewRow(sheet, d, '⚠️ 注意：第 ' + exact[0] + ' 列已有同規格但承攬商為「' +
      (existAgent || '未填') + '」的紀錄，本次已以「' + (newAgent || '未填') + '」另外加一列');
  }
  return appendNewRow(sheet, d, '（依「強制新增」直接入表）');
}

function appendNewRow(sheet, d, note) {
  var newRow = buildRowValues(d);
  sheet.appendRow(newRow);
  var out = '【🟢 新增成功】已寫入第 ' + sheet.getLastRow() + ' 列\n' + formatRecordLine(newRow, '');
  return note ? out + '\n' + note : out;
}

// ------------------ 2. 修改 ------------------
function doUpdate(sheet, d, originalText) {
  var C = CONFIG.COLUMNS;
  if (!d.temple && !d.agent) {
    return "【⚠️ 修改失敗】請至少提供「廟宇」名稱，例：修改 新化武廟 4*5 OLED琥珀色 總燈數變成 5238 盞";
  }

  var data = readData(sheet);
  // 「某某 業務改成 聖文」這類是改名，業務欄不可拿來當定位條件
  var agentRename = /業務\s*(?:改成|換成|改為|更新為)|(?:改成|換成)業務/.test(originalText || '');
  var collect = function(useSpec) {
    var out = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (isSummaryRow(row)) continue;
      if (d.temple && !matchTemple(cell(row, C.TEMPLE), d.temple)) continue;
      if (d.agent && !agentRename && !containsLoose(cell(row, C.AGENT), d.agent)) continue;
      if (useSpec && d.spec && !matchSpec(cell(row, C.SPEC), d.spec)) continue;
      out.push(i + 1);
    }
    return out;
  };

  var hits = collect(true);
  // 「規格改成 ...」這類輸入：找不到同規格列時，退而用廟名定位（僅限該廟只有一列），
  // 才允許換規格欄，避免同廟多規格時改錯列
  var specChange = false;
  if (hits.length === 0 && d.spec && /規格\s*[改換調]|改成|換成|改為/.test(originalText || '')) {
    var alt = collect(false);
    if (alt.length === 1) { hits = alt; specChange = true; }
  }

  if (hits.length === 0) {
    var tips = [];
    for (var k = 1; k < data.length; k++) {
      if (isSummaryRow(data[k])) continue;
      if (d.temple && matchTemple(cell(data[k], C.TEMPLE), d.temple)) {
        tips.push('    · 第' + (k + 1) + '列 ' + (cell(data[k], C.SPEC) || '未填規格'));
      }
      if (tips.length >= CONFIG.MAX_LIST_ROWS) break;
    }
    return "【❌ 修改失敗】找不到「" + (d.temple || d.agent) + "」" + (d.spec ? ' ＋規格「' + d.spec + '」' : '') + ' 的資料。' +
           (tips.length ? "\n📋 該廟既有規格：\n" + tips.join('\n') : '') +
           "\n💡 輸入 /help 查看使用方式";
  }

  // 同廟多規格：未指明規格時先請使用者補充，避免改錯列
  if (hits.length > 1) {
    var opts = hits.slice(0, CONFIG.MAX_LIST_ROWS).map(function(r) {
      return '    · 第' + r + '列 ' + (cell(data[r - 1], C.SPEC) || '未填規格') + '（' + withComma(cell(data[r - 1], C.TOTAL)) + ' 盞）';
    });
    return "【⚠️ 需要指明規格】「" + (d.temple || d.agent) + '」符合 ' + hits.length + " 筆：\n" + opts.join('\n') +
           (d.spec ? "\n（同一規格也有多筆，請再指明業務，或改用「強制新增」補一筆）" : '') +
           '\n例：修改 ' + (d.temple || '') + ' ' + (cell(data[hits[0] - 1], C.SPEC) || '5*7 OLED琥珀色') + ' 總燈數變成 100 盞';
  }

  var r = hits[0];
  var target = data[r - 1];
  var changes = [];
  var apply = function(colEnum, label, newVal, opts) {
    if (newVal === null || newVal === undefined || newVal === '') return;
    var fmt = (opts && opts.fmt) || function(v) { return v; };
    var unit = (opts && opts.unit) || '';
    var oldVal = cell(target, colEnum);
    if (String(oldVal) === String(newVal)) return;
    sheet.getRange(r, colEnum).setValue(newVal);
    changes.push('    ' + label + '：' + (oldVal === '' ? '（原空白）' : fmt(oldVal) + unit) + ' → ' + fmt(newVal) + unit);
  };

  apply(C.AGENT, '業務', d.agent);
  apply(C.TOTAL, '總燈數', d.total_count ? toNumber(d.total_count) : null, { fmt: withComma, unit: ' 盞' });
  if (d.delivery_text) apply(C.DELIVERY, '送燈時間', d.delivery_text);
  if (d.delivery_text || d.remark) {
    var nextDelivery = d.delivery_text || cell(target, C.DELIVERY);
    var nextRemark = d.remark !== null && d.remark !== undefined && String(d.remark).trim() !== ''
      ? d.remark : stripDateTag(cell(target, C.REMARK));
    apply(C.REMARK, '備註', buildRemark(nextRemark, nextDelivery));
  }
  apply(C.SOFTWARE, '軟體', d.software);
  apply(C.COMPUTER, '電腦', d.computer);

  apply(C.SPEC, '規格', specChange ? d.spec : null);

  var title = "【🟡 修改成功】第 " + r + " 列｜" + (cell(target, C.AGENT) || '未註明') + '－' + (cell(target, C.TEMPLE) || '未註明') +
              '（' + (cell(target, C.SPEC) || '未填規格') + '）';
  if (changes.length === 0) return title + "\n    ℹ️ 提供的欄位與現有數值相同，未做更新。";
  return title + '\n' + changes.join('\n');
}

// 批次把業務／公司名改名（GAS 編輯器執行；dryRun=true 只預覽不寫入）
// 例：renameAgent('明典', '聖文', true) 先看會改到哪些列，確認後再改 renameAgent('明典', '聖文')
function renameAgent(fromName, toName, dryRun) {
  var C = CONFIG.COLUMNS;
  var sheet = getLightSheet();
  if (!sheet) return '找不到分頁「' + CONFIG.SHEET_NAMES.LIGHT_MGMT + '」，請先執行 setupLightSheet()';
  if (!fromName || !toName) return '請輸入兩個名稱：renameAgent("舊名", "新名")';

  var data = readData(sheet), rows = [], preview = [];
  for (var i = 1; i < data.length; i++) {
    if (isSummaryRow(data[i])) continue;
    var cur = cell(data[i], C.AGENT);
    if (!cur || cur.indexOf(fromName) === -1) continue;
    var next = cur.split(fromName).join(toName);
    rows.push(i + 1);
    preview.push('    第 ' + (i + 1) + ' 列：' + cur + ' → ' + next + '（' + cell(data[i], C.TEMPLE) + '）');
    if (!dryRun) sheet.getRange(i + 1, C.AGENT).setValue(next);
  }

  // 使用者綁定的身分也一起換掉，免得之後又寫回舊名
  var bindHit = 0;
  if (!dryRun) {
    try {
      var map = readIdentities(), changed = false;
      Object.keys(map).forEach(function(uid) {
        var idt = map[uid] || {};
        ['company', 'person', 'label'].forEach(function(f) {
          if (idt[f] && String(idt[f]).indexOf(fromName) !== -1) { idt[f] = String(idt[f]).split(fromName).join(toName); changed = true; bindHit++; }
        });
      });
      if (changed) PropertiesService.getScriptProperties().setProperty(IDENTITY_KEY, JSON.stringify(map));
    } catch (e) { console.error('renameAgent 更新綁定失敗: ' + e.toString()); }
  }

  var msg = (dryRun ? '【預覽】' : '【已完成】') + '「' + fromName + '」→「' + toName + '」共 ' + rows.length + ' 列' +
            (bindHit ? '；另更新 ' + bindHit + ' 個已綁定身分' : '') + '\n' + (preview.join('\n') || '    （沒有符合的列）');
  Logger.log(msg);
  return msg;
}

// ------------------ 3. 查詢 ------------------
function doRead(sheet, d, originalText) {
  var C = CONFIG.COLUMNS;
  var data = readData(sheet);
  var hasFilter = !!(d.agent || d.temple || d.spec || d.software || d.computer);
  var hits = [];
  var totalSum = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (isSummaryRow(row)) continue;
    if (!hasFilter) { hits.push(row); continue; }
    var ok = true;
    if (d.temple && !matchTemple(cell(row, C.TEMPLE), d.temple)) ok = false;
    if (ok && d.spec && !matchSpec(cell(row, C.SPEC), d.spec)) ok = false;
    if (ok && d.agent && !containsLoose(cell(row, C.AGENT), d.agent)) ok = false;
    if (ok && d.software && !containsLoose(cell(row, C.SOFTWARE), d.software)) ok = false;
    if (ok && d.computer && !containsLoose(cell(row, C.COMPUTER), d.computer)) ok = false;
    if (ok) hits.push(row);
  }

  if (hits.length === 0) {
    return "【📭 查詢無結果】找不到符合條件的燈位紀錄。\n💡 可試：查一下 金六結福德廟 ／ 查詢 5*7 OLED ／ 查 冠緯 的軟體燈位";
  }

  var matched = hasFilter ? hits.slice(0, CONFIG.MAX_LIST_ROWS) : hits.slice(-CONFIG.MAX_LIST_ROWS);
  hits.forEach(function(row) { totalSum += toNumber(cell(row, C.TOTAL)); });

  var lines = matched.map(function(row) { return formatRecordLine(row, ''); });
  var head = hasFilter
    ? '【🔍 查詢結果】共 ' + hits.length + ' 筆｜合計 ' + withComma(totalSum) + ' 盞'
    : '【🔍 最近 ' + matched.length + ' 筆登記】共 ' + hits.length + " 筆資料（請指定廟宇或規格可縮小範圍）";
  var tail = hits.length > matched.length ? "\n……（" + (hasFilter ? '僅顯示前 ' : '僅顯示最新 ') + matched.length + ' 筆，共 ' + hits.length + ' 筆）' : '';
  return head + '\n\n' + lines.join('\n------------------\n') + tail;
}

// ==================== 5. LINE 回傳訊息工具（見 5.5 送出層） ====================
// ==================== 5.5 LINE 送出層（單則 5000 字上限、一次最多 5 則） ====================
var LINE_TEXT_LIMIT = 4800;   // 官方單則 text 上限 5000 字元，留 200 字安全餘量
var LINE_MAX_MESSAGES = 5;    // 官方：一次 reply 最多 5 個 message objects

// 依換行切段，切點太爛就硬切，避免把一個字元/半個網址切開
function splitLineText(text, size) {
  var s = String(text === null || text === undefined ? '' : text).replace(/\r/g, '');
  var out = [];
  while (s.length > size) {
    var cut = s.lastIndexOf('\n', size);
    if (cut < Math.floor(size * 0.5)) cut = size;
    out.push(s.substring(0, cut));
    s = s.substring(cut).replace(/^\n+/, '');
  }
  if (s.length) out.push(s);
  return out.length ? out : [''];
}

// 超長內容：先切則，超過 5 則就把多出的部分省略並告知
function prepareLineMessages(messageText) {
  var chunks = splitLineText(messageText, LINE_TEXT_LIMIT);
  if (chunks.length <= LINE_MAX_MESSAGES) return chunks;

  var kept = chunks.slice(0, LINE_MAX_MESSAGES - 1);
  var dropped = chunks.slice(LINE_MAX_MESSAGES - 1).join('\n');
  kept.push(dropped.substring(0, LINE_TEXT_LIMIT - 120) +
            '\n……（內容過長，已省略約 ' + (dropped.length - (LINE_TEXT_LIMIT - 120)) + ' 字；請改用瀏覽器看試算表，或縮小問題再問一次）');
  return kept;
}

function sendLineReply(replyToken, messageText) {
  // ⭐ 必須是這個完整的官方 API 網址，LINE 才能收到你的回信
  var url = "https://api.line.me/v2/bot/message/reply";
  var messages = prepareLineMessages(messageText).map(function (t) {
    return { "type": "text", "text": t };
  });

  var payload = { "replyToken": replyToken, "messages": messages };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": "Bearer " + CONFIG.LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true // 被 LINE 拒絕時 GAS 不會崩潰，但一定要把狀態碼讀出來報錯
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode(), body = response.getContentText();
  console.log('LINE 回應狀態碼: ' + code + '（送出 ' + messages.length + ' 則、' + String(messageText || '').length + ' 字）');
  if (code !== 200) {
    // 以前這裡只 log，結果超長或被拒時使用者只看到「沒回訊」，現在把原因丟回去
    console.error('LINE 送出失敗: ' + body);
    throw new Error('LINE 送出失敗 HTTP ' + code + '：' + String(body).substring(0, 200));
  }
  return messages.length;
}

// ==================== 6. 網頁瀏覽器直接開啟端 (GET) ====================
function doGet(e) {
  // 當用瀏覽器直接點開網址時，會看到這個漂亮精簡的網頁回應
  var htmlContent = 
    "<div style='font-family: Arial, sans-serif; text-align: center; margin-top: 100px; padding: 20px; border: 1px solid #ddd; display: inline-block; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>" +
    "  <h1 style='color: #d9534f; margin-bottom: 5px;'>⛩️ 宮廟光明燈管理系統</h1>" +
    "  <p style='color: #666; font-size: 14px;'>Google Apps Script 後端雲端服務</p>" +
    "  <hr style='border: 0; border-top: 1px solid #eee; margin: 20px 0;'>" +
    "  <div style='background-color: #dff0d8; color: #3c763d; padding: 10px 20px; border-radius: 4px; font-weight: bold;'>🟢 系統連線狀態：正常運作中</div>" +
    "  <p style='color: #888; font-size: 12px; margin-top: 20px;'>請透過已綁定的 LINE 官方帳號進行「新增、修改、查詢」操作；輸入 /help 可查看使用方式。</p>" +
    "</div>";
    
  return HtmlService.createHtmlOutput(htmlContent).setTitle("光明燈管理系統端點");
}

function testGeminiProbe() {
  var candidates = getCandidateModels();
  Logger.log('📋 金鑰可用模型: ' + candidates.join(' , '));
  var r = analyzeMessageWithGemini('聖文-石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前 軟體其他 電腦研華');
  Logger.log('✅ Gemini 解析成功: ' + JSON.stringify(r));
  Logger.log('📅 輔助日期欄將寫入: ' + (parseDateKey(r.details.delivery_text) || '(空白)'));
}

function testLineToken() {
  var resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
    headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_ACCESS_TOKEN },
    muteHttpExceptions: true
  });
  Logger.log('🔑 LINE Token 檢查 HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText());
}
