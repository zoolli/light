# ⛩️ 宮廟光明燈管理系統 — 建構與規格文件

一支 Google Apps Script（以下 GAS）同時當作 LINE Bot 後端與試算表寫入端：
經辦人在 LINE 用**自然語言**登記燈位，AI 解析後寫進 Google 試算表；查詢與大量修改直接用瀏覽器開試算表。

- 程式碼：`light.gs`（單檔，無相依套件）
- 測試：`test/light.test.js`（純 Node 跑的 stub 測試，不需 GAS 環境）→ `node test/light.test.js`，目前 101 項全過

---

## 1. 系統規格（架構）

```
經辦人手機 LINE
      │ 文字訊息（需帶喚醒字「小幫手」）
      ▼
LINE 官方帳號（Messaging API Webhook）
      │ POST JSON
      ▼
GAS Web 應用程式  doPost(e)
      ├─ 1 指令分流：/help、/model（不打 AI）
      ├─ 2 回覆策略：未帶喚醒字／查詢／閒聊 → 沉默（連 AI 都不調，省配額）
      ├─ 3 身分綁定：「我公司 亞盛燈業 我叫 李小華」→ 存腳本屬性 USER_IDENTS
      ├─ 4 Gemini generateContent（模型序：手動鎖定 > 預設 > 自動探測，最多 6 個）
      │      回傳 JSON：{ action, details{...} }
      ├─ 5 必填攔截：缺欄 → 存 10 分鐘草稿 + 回覆「還需要什麼」與可照抄的格式
      └─ 6 試算表寫入／修改（A~H 欄；跳過合計列）
      │
      ▼
Google 試算表「光明燈管理」分頁  ←→  瀏覽器直接編輯（查詢、小計公式都在這邊）
      │
      ▼
sendLineReply() → 手機收到「新增成功／修改成功／需要補件」的回覆
```

| 元件 | 用途 | 備註 |
|---|---|---|
| LINE Messaging API | 收發訊息 | 只能拿到 `userId`；`displayName` 要另外呼叫 profile API |
| Google Apps Script | Webhook 主機 + 業務邏輯 | 免伺服器，免費額度內 |
| Gemini API | 中文語意 → JSON | `responseMimeType: application/json` |
| Google 試算表 | 唯一資料來源 | 經辦人可用瀏覽器像 Excel 一樣直接編輯 |

---

## 2. 資料規格（試算表欄位）

分頁名稱：`光明燈管理`（`CONFIG.SHEET_NAMES.LIGHT_MGMT`）。第 1 列是表頭，可由 `setupLightHeader()` 自動寫入。

| 欄 | 欄位 | 必填 | 格式／範例 | 說明 |
|---|---|---|---|---|
| A | 業務 | ✅ | `聖文`、`亞盛燈業-李小華` | 公司名、經辦人，或「公司-人員」。未填時套用該 LINE 使用者綁定的身分 |
| B | 廟宇 | ✅ | `石岡子乾元宮`、`天成宮(北投)` | 比對時忽略空白、全形括號、頓號 |
| C | 規格 | ✅ | `5*7 OLED琥珀色`、`4*5 OLED` | 尺寸＋燈種＋顏色，**原樣保留** |
| D | 總燈數 | ✅ | `2112` | 純數字；輸入 `2,112盞` 會剝成 2112 |
| E | 送燈時間 | － | `國10/17前`、`12/10or12/17`、`已送燈` | **保留使用者的寫法**，不被系統改寫 |
| F | 軟體 | － | `廟幫手`、`冠緯`、`其他` | 可省略，之後用瀏覽器補 |
| G | 電腦 | － | `研華`、`廟`、`研華*3` | 同上 |
| H | 備註 | － | `〔國2026-10-17〕 分兩批送，第二批國12/01前` | **其他資訊全放這**。訊息裡明確寫「備註：…」才會填入說明；系統另外把 E 欄推得出的國曆日期以〔國…〕標記前置 |
| I→ | 你的公式欄 | － | `=SUM(D:D)` | 合計、小計請放 **I 欄以後**，系統只看 A~H |

### 2.0 主表與公司分頁（讀多寫一）

`CONFIG.SHEET_NAMES` 除了主表还可以列各承攬商分頁，但它們**只是主表的唯讀投影**，不是另一份資料：

```javascript
SHEET_NAMES: {
  LIGHT_MGMT: '光明燈管理',   // 主表：機器人只寫這裡，也是全部資料
  SHENG_WEN:  '聖文',        // 唯讀投影：QUERY 自主表過濾 A 欄
  MING_DIAN:  '明典'         // 同上
}
```

- 執行 `setupLightSheet()` 會一次完成：建主表＋表頭＋凍結，再建/更新各公司分頁（A1 表頭、A2 起 `=QUERY('光明燈管理'!A2:H, "where upper(A) contains upper("聖文")", 0)`），並把投影頁設成「編輯時警告」
- **没列進去的公司（例如冠緯）就只留在主表**，不會多出分頁，也不影響合計
- 公司分頁名稱可自由改（值改變就好），KEY 請維持大寫蛇標式
- 投影範圍是主表的 `A2:H`（含備註欄）；欄位以後再加，投影會自動跟著長寬走（`maxColumnIndex()`）
- 不需要投影機制就把那兩行 KEY 刪掉，`syncCompanyTabs()` 會直接回傳「未設定公司分頁」，整個行為回到單表
- 修改/新增一律走主表；投影頁因為是公式，不允許也寫不進去（別試圖在 A2 打字，會把整個投影蓋掉）

### 2.1 規格（C 欄）比對正規化
小寫、去空白、去掉 `* × x ✕ ＊` 之後再比，並允許互相包含：

| 你輸入 | 系統視為 |
|---|---|
| `5*7 OLED琥珀色` / `5×7 oled 琥珀色` / `5 7 OLED琥珀色` | 同一筆 |
| `4*5 OLED` | **不等於** `5*7 OLED` |
| `5*7` | 可對到 `5*7 OLED琥珀色`（較短者包含較長） |

> 同廟常有 4*5 與 5*7 兩列，所以修改時**不寫規格**又同時符合兩列，系統不會亂改，會列出既有規格請你指明。

### 2.2 送燈時間：國曆／民國／農曆 對應表（E 原樣 → 備註欄的〔國…〕標記）

系統**只換算國曆**。E 欄永遠保留你打的字，換算出的國曆日期以 `〔國YYYY-MM-DD〕` 前置進 H 備註欄，方便排序與篩選。

| E 欄原樣 | 備註欄標記 | 說明 |
|---|---|---|
| `國10/17前`、`國曆10/17前` | `2026-10-17` | 「國」＝國曆；只有月日時補今年 |
| `10月15日`、`10/15` | `2026-10-15` | 只有月日時補今年（不跨年推估） |
| `國115/02/09前`、`國115-02/09前`、`民國115/02/09` | `2026-02-09` | 2~3 位數年份一律當民國：`115 + 1911` |
| `國116/01/02` | `2027-01-02` | 同上 |
| `西元2027-06-30`、`2026-11-07` | 原樣 | 4 位數年份直接當西元 |
| `12/10or12/17` | `2026-12-10` | 兩個日期取第一個（E 欄仍留 `12/10or12/17`） |
| `已送燈`、`送完`、`待定`、空白 | 留空 | 沒有任何日期可推 |
| `農曆10/17前`、`舊曆11/05`、`陰曆12/01前` | **不插標記（刻意不推）** | GAS 沒有農曆⇄國曆對照表，硬換算一定錯。要查請自己對萬年曆，或在 E 欄直接補國曆 |
| `國10/32`、月份 > 12 等明顯錯值 | 留空 | 防呆，不會吐出怪日期 |

> 需要「農曆也自動換算」的話，得內掛一張農曆對照表（約 60 個年份、一萬多字），我可以另外做，但那會把 `light.gs` 撐大一大截 —— 目前選擇不猜。
> 不想在備註欄看到標記：腳本屬性 `DATE_TAG_IN_REMARK = OFF`（備註欄就只放你自己寫的說明）。
> 標記是**冪等**的：重複修改備註不會堆疊，換了送燈時間會把舊標記換掉，自己打的說明一律照原文保留。

### 2.3 寫入與保護規則
- 新增一律 `appendRow` 加在**最後一列**。小計若寫死範圍 `=SUM(D2:D20)` 會漏計，請改 `=SUM(D:D)` 或 `=SUM(D2:D)`
- 含有 `合計／總計／小計` 字樣的列（只看 A~C 欄）一律跳過，不會被查詢列出、也不會被修改覆寫
- 同廟同規格已存在時，新增會被擋下並顯示既有那一列；真的是第二批請在訊息加 `強制新增`

---

## 3. 行為規格（LINE 端）

處理順序（`decideFlow()` → AI → `shouldReplyFor()`）：

1. `/help`、`/model` → **一定回覆**
2. 身分綁定句（整句只有關鍵字＋名稱）→ 記住並回覆，不動資料
3. **沒帶喚醒字** → 完全沉默（不打 AI，省配額）。例外：該使用者有 10 分鐘內的補件草稿，或 `REPLY_MODE=ALL`
4. 查詢語氣（`查一下/有哪些/幾盞/總共…` 且無寫入動詞）→ 沉默，請用瀏覽器看試算表
5. 打 AI 後只有 `CREATE`／`UPDATE` 會回訊；`READ`／無法辨識 → 沉默並寫執行記錄

### 3.1 必填與補件
`REQUIRED_FIELDS = 業務、廟宇、規格、總燈數`。缺任何一項**完全不上表**，回覆：

```
【🟡 資料還沒齊全，暫未登記】
📥 已收到：廟宇 媽祖廟、規格 財神燈
❓ 還需要：業務／公司名、總燈數

✍️ 可以直接補：小幫手 聖文 2112盞
　（業務只需綁定一次：「我是 聖文」或「我公司 亞盛燈業 我叫 李小華」）
⏳ 10 分鐘內補件我會自動合併，不必重打整串；若想先佔位之後再補，請加「直接新增」
```

- 草稿存在 `CacheService`，TTL 600 秒（`CONFIG.DRAFT_TTL`）。補件時系統會把「上一句＋這一句」串起來再送 AI，所以只回 `500盞` 也接得上
- 想先佔位再加「直接新增」；想改業務欄請說「業務改成 聖文」

### 3.2 業務身分綁定（公司／派來的人員）
整句只有關鍵字與名稱才會被當成綁定（句中有殘字一律視為登記內容，避免吞掉資料）：

| 你輸入 | A 欄日後自動帶入 |
|---|---|
| `小幫手 我是聖文` | `聖文` |
| `小幫手 我公司 亞盛燈業 我叫 李小華` | `亞盛燈業-李小華` |
| `小幫手 我是 亞盛燈業 的 李小華` | `亞盛燈業-李小華` |
| `小幫手 我叫 李小華` | `李小華` |
| `小幫手 業務窗口 冠緯` | `冠緯` |

**A 欄（業務）的取值優先序** —— 本系統預設由「聖文」承攬，所以沒特別交代就上聖文：

| 優先序 | 來源 | 回覆會顯示 |
|---|---|---|
| 1 | 句中寫的（`名典-○○宮 …`） | 不帶入提示，直接用你寫的 |
| 2 | 該 LINE 使用者綁定的公司／人員 | `👤 業務自動帶入（您的綁定）：名典-李小華` |
| 3 | `DEFAULT_AGENT`（預設 `聖文`） | `👤 業務預設帶入：聖文（若是别家訂單請寫在句首…）` |
| 4 | LINE 暱稱（需 `AGENT_FROM_LINE_PROFILE=ON` 且 3 設空） | `👤 業務自動帶入 LINE 暱稱「…」` |
| － | 上述全無 | 不回問？會 → `還需要：業務／公司名` |

- 換人／換公司：再傳一次會蓋掉；單筆想填別家，直接寫在句首（`名典-○○宮 …`）
- 整組訂單都是别家承包：把腳本屬性 `DEFAULT_AGENT` 改成那家名稱即可，不用改程式
- **預設不會**拿 LINE 暱稱寫進業務欄（暱稱常是個人綽號，會弄髒公司名）。需要時在腳本屬性加 `AGENT_FROM_LINE_PROFILE = ON`

### 3.3 指令規格

| 指令 | 作用 |
|---|---|
| `/help`（或 `使用方式`、`說明`、`?`） | 回完整使用說明（約 1000 字，單則訊息） |
| `/model` | 看目前模型與這把金鑰可用的 flash 模型清單（快取 6 小時） |
| `/model gemini-3.1-flash-lite` | 手動鎖定模型 |
| `/model auto` | 取消鎖定，回到自動依序嘗試 |

### 3.4 可改的預設值

| 位置 | 預設 | 說明 |
|---|---|---|
| `CONFIG.WAKE_WORDS` | `小幫手, 小帮, 幫手, 助理` | 腳本屬性 `WAKE_WORDS`（逗號分隔）可蓋，免改程式 |
| `CONFIG.REQUIRED_FIELDS` | `agent, temple, spec, total_count` | 必填欄位（agent 會由優先序自動補齊） |
| `CONFIG.DEFAULT_AGENT` | `聖文` | 業務欄的預設承攬商，腳本屬性可蓋 |
| `CONFIG.DRAFT_TTL` | `600` 秒 | 補件草稿保留；CacheService 上限 600 |
| `CONFIG.MAX_LIST_ROWS` | `12` | 查詢／提示最多列出筆數 |
| `CONFIG.COLUMNS` | A~H | 欄位索引對應表，改表結構只動這裡 |

---

## 4. 建置步驟

### 4.1 建立 LINE 官方帳號（約 10 分鐘）
1. [LINE Official Account Manager](https://manager.line.biz) 建立帳號 → 設定 → **MESSENGER API** → 「使用 LINE Messaging API」→ 選擇/建立 Provider，取得 **Channel ID**
2. [LINE Developers Console](https://developers.line.biz/console/) → 該 Channel → **Messaging API** 頁 → `Issue`（選 **channel access token long-lived**）→ 複製 token
3. Official Account Manager → 設定 → **MESSENGER API** → 開啟「允許來自 LINE 平台的存取」（關著的話 Webhook 永遠收不到資料）
4. Official Account Manager → 設定 → **回應設定** → 關閉「自動回應訊息」（避免與_bot_回覆重複）
5. 若要使用 LINE 暱稱自動帶入（`AGENT_FROM_LINE_PROFILE`）不需另外申請 LINE Login：profile API 用的就是第 2 步那組 channel access token

### 4.2 建立 GAS 專案
1. <https://script.google.com> → 新的專案 → 把 `light.gs` 整份貼上並儲存
2. 專案設定（⚙️）→ **Script Properties** → 新增：

| 屬性 | 值 | 必要性 |
|---|---|---|
| `LINE_ACCESS_TOKEN` | 4.1 的 long-lived token | 必要 |
| `GEMINI_API_KEY` | Google AI Studio 的金鑰 | 必要 |
| `SPREADSHEET_ID` | 試算表網址 `/d/` 與 `/edit` 之間那串 | 建議必填，詳見 4.3 |
| `GEMINI_MODEL` | 例 `gemini-3.1-flash-lite` | 選填，等同 `/model xxx` |
| `WAKE_WORDS` | 例 `小幫手,助理` | 選填 |
| `SHEET_NAME` | 主表分頁名稱，預設 `光明燈管理` | 選填（分頁由 `setupLightSheet()` 自動建立） |
| `DEFAULT_AGENT` | 沒寫業務、也沒綁定時的預設承攬商，預設 `聖文`；設空白＝改用綁定／LINE 暱稱，都沒有就回問 | 選填 |
| `REPLY_MODE` | `ALL` 時查詢與閒聊也會回訊 | 選填 |
| `DATE_TAG_IN_REMARK` | 設 `OFF` 就不把〔國…〕標記寫進備註欄 | 選填 |
| `AGENT_FROM_LINE_PROFILE` | `ON` 時未填業務自動帶 LINE 暱稱 | 選填 |

3. 分頁名稱改成 `光明燈管理`，並把 A 欄拆成「業務／廟宇」兩欄（原有資料整欄右移），合計公式挪到 I 欄以後
4. 選單函式選 `setupLightHeader` → 執行（首次會要求授權 Google 帳戶，同意 Sheets 權限）→ 表頭 A1:H1 完成
5. 可先跑 `testLineToken()`、`testGeminiProbe()` 確認兩條金鑰都通（看「執行記錄」）

### 4.3 試算表 ID 與權限（資料究竟落在哪一張表）

`CONFIG.SPREADSHEET_ID`（light.gs:7）依下面三個順序解析，第一筆拿到就用它：

| 順序 | 來源 | 適用情境 |
|---|---|---|
| 1 | 腳本屬性 `SPREADSHEET_ID` | 獨立 GAS 專案（本文件走这条路）。**建議一律用它**，最明確也好搬表 |
| 2 | `SpreadsheetApp.getActiveSpreadsheet().getId()` | 用「試算表 → 扩展程序 → Apps Script」建立的**綁定型腳本**，會自動拿到它所在的那張表 |
| 3 | 硬編 fallback ID `1Kq6Du15…TH9A` | 前兩者都拿不到時的防呆。**換表時最容易忘記改，建議確認 1 或 2 生效後刪掉這段** |

ID 怎麼看：試算表網址 `https://docs.google.com/spreadsheets/d/←這一串→/edit`

換表／搬資料流程（四步）：
1. 新表建好，分頁名稱取為 `光明燈管理`，舊資料整欄貼上（業務與廟宇記得拆成 A、B 兩欄）
2. 腳本屬性 `SPREADSHEET_ID` 改成新 ID（綁定型腳本可整個跳過 1~2，直接改分頁名就好）
3. 執行一次 `setupLightHeader()` → 確認 A1:H1 表頭正確、欄數補到 H
4. 傳一筆測試登記 → 再到瀏覽器看該列 → 沒問題才正式開始；備註：`doRead` 只讀不寫，可當驗證工具

權限要注意的事：

- 部署時選「執行對象：**我**」→ 所有讀寫都用**你這個 Google 帳號**的權限。經辦人只需要是 LINE 好友，**不需要** Google 帳號、也**不需要**把試算表分享給他
- 所以這張表必須在你自己帳號底下（或至少你是編輯者），否則 `openById()` 會直接報權限錯誤
- 千萬不要為了經辦人改成「執行對象：任何使用者」，那樣每位經辦人都要有自己的 Google 帳號 + 這張表的編輯權限，等於用回舊方法
- 檔案裡 `getLightSheet()` 用的是「分頁名稱」（光明燈管理）而不是分頁索引，所以分頁順序可以亂放，但**名稱不能改**

### 4.4 部署並接上 Webhook
1. 右上 **部署 → 新的部署項目** → 類型選「網頁應用程式」
   - 執行對象：**我**
   - 誰可以存取：**任何人**（若只想給自家經辦人，改「任何-google.com 帳號」並在試算表共用给他们）
2. 複製產生的網址（結尾 `/exec`）→ 貼到 LINE Developers Console → **Webhook URL** → **Verify**（應顯示成功）→ 打開 **Use webhook** → 儲存
3. 之後每次改程式碼都要：**部署 → 管理部署 → 編輯 → 版本「新版本」→ 部署**（網址不变）。改完沒發新版本是最常見的「沒反應」原因

### 4.5 上線後第一件事
```
小幫手 石岡子乾元宮 5*7 OLED琥珀色 2112盞 國10/17前   ← 沒寫業務就上預設「聖文」
小幫手 修改 新化武廟 4*5 OLED琥珀色 總燈數變成5300盞
小幫手 我是 名典                                    ← 只有别家的人／窗口要綁一次
```

---

## 5. 給 AI 的契約（prompt 規格）

`analyzeMessageWithGemini()` 送出的 JSON 固定為：

```json
{ "action": "CREATE | UPDATE | READ",
  "details": { "agent": "", "temple": "", "spec": "", "total_count": 0,
               "delivery_text": "", "software": "", "computer": "" } }
```

prompt 內的四條硬規則（要加欄位時一起改）：
- A. `聖文-石岡子乾元宮` 這種連寫要拆成 `agent` + `temple`
- B. 日期一律**照抄**到 `delivery_text`，不換算 ISO（換算由 GAS 的 `parseDateKey()` 負責）
- C. 同廟可能有多規格，`spec` 有提到就必須填
- D. 語意判斷：登記類 `CREATE`、改成/調整 `UPDATE`、查詢 `READ`

模型失敗處理：429／500／503／504 會自動換下一個模型（最多 6 個，間隔 1.2 秒），全掛時回傳中文提示並建議 `/model gemini-3.1-flash-lite`。

---

## 6. 排錯對照表

| 狀況 | 原因／檢查點 | 處置 |
|---|---|---|
| 訊息完全沒回 | 沒帶喚醒字、或那是查詢句（設計如此） | 執行記錄會印 `🔇 不回覆（…）`；要全部回覆就設 `REPLY_MODE=ALL` |
| LINE Verify 失敗 | 沒發新版本、網址少 `/exec`、權限不是「任何人」 | 重走 4.4 |
| `找不到分頁「光明燈管理」` | 分頁還沒建立或名稱不對 | GAS 執行一次 `setupLightSheet()`（會自動建分頁），或用腳本屬性 `SHEET_NAME` 指定現有分頁名 |
| 只有「系統異常」 | 試算表分頁名不對、`SPREADSHEET_ID` 錯、該表不在你帳號下 | 看執行記錄堆疊；`找不到名為「光明燈管理」的工作表分頁` = ID 對但分頁名錯；`openById` 權限錯誤 = 表不屬於你 |
| 登記寫到舊表 | 拿到的是 light.gs:7 的 fallback ID（腳本屬性沒設、又不是綁定型腳本） | 設 `SPREADSHEET_ID` 或把 fallback 那段刪掉 |
| `AI 免費配額暫時用盡` | Gemini 免費層限流 | `小幫手` 不用管，稍後重試；或 `/model gemini-3.1-flash-lite` |
| 業務欄變成綽號 | 開了 `AGENT_FROM_LINE_PROFILE` | 傳 `我公司 ○○` 綁定，或關掉該屬性 |
| 新增後小計沒變 | 新增列落在小計**下方** | 合計改整欄 `=SUM(D:D)`，或把小計移到 I 欄以後 |
| 改到不該改的列 | 同廟多規格未指明 | 系統會先列清單要你補規格，不會自行猜 |
| 拿不到 LINE 姓名（404） | 對方沒加為好友／封鎖／未同意提供 profile | 屬正常，改用「我是 ○○」綁定 |

---

## 7. 測試

```bash
node test/light.test.js
```
用 Node 的 `vm` 载入 `light.gs`，把 `SpreadsheetApp`／`UrlFetchApp`／`CacheService` 等 GAS 服務换成 stub，並塞入真實格式的 17 列種子資料（含合計列、右側公式欄）。涵蓋（共 101 項）：

- 日期正規化（國曆／民國／農曆不換算）十餘種寫法
- 備註欄：明示才寫、可單獨修改、查詢會列出；〔國…〕標記的插入／替換／不堆疊、可關掉
- 小計列跳過、規格／廟名正規化比對
- 新增：重複攔截、`強制新增`、8 欄落地
- 修改：只改講到的欄位、多規格先問、`業務改成`、合計列不可動
- 查詢：筆數與合計盞數、無條件時只列最新且不含合計
- 回覆策略：喚醒字、查詢／閒聊沉默、`REPLY_MODE=ALL`
- 端到端 `doPost`：缺欄追問 → 10 分鐘內補件合併 → 寫入
- 業務欄優先序（句中 > 綁定 > `DEFAULT_AGENT` > 暱稱）
- 分頁自動建立 `setupLightSheet()`、`SPREADSHEET_ID` 佔位值防護
- 公司分頁只產 QUERY 投影、不為其他家亂開分頁、拿掉 `SHEET_NAMES` 兩行可整組停用
- 身分綁定各種寫法、批次改名 `renameAgent()` 預覽與實作

改動 `light.gs` 後請先跑这支再部署。

---

## 8. 已知極限與注意事項

- 單檔 GAS：執行時間上限 6 分鐘（本系統單次需求 <2 秒）；`PropertiesService` 内容 500KB、`USER_IDENTS` 程式內只保留最近 300 位使用者
- 同一人 10 秒內連發兩則可能抓到舊草稿（webhook 無鎖），建議逐則送出、收到回覆再傳下一筆
- LINE 單則文字 5000 字上限；`/help` 約 1036 字，查詢結果超過 12 筆會截斷
- 檔案後半段 `testAllPermissions`、`testCreateCalendarEvent`、`testCreateContact`、`createGoogleCalendarEvent`、`createGoogleContact`（約 1039~1201 行）是另一專案（教師管家）的殘留，與光明燈無關，但會讓**首次授權多要求 Google 日曆與聯絡人權限**。不需要就直接把那幾支函式整段刪除，不影響本系統

---

## 9. 程式碼地圖（`light.gs` 區塊與職責）

| 區塊 | 內容 | 要改行為時看這裡 |
|---|---|---|
| `1. 設定區` | `CONFIG`：金鑰、分頁名、`COLUMNS`、`HEADERS`、`WAKE_WORDS`、`REQUIRED_FIELDS`、`FIELD_LABELS`、`DRAFT_TTL` | 欄位順序、必填、喚醒字 |
| `2. Webhook 接收端` | `doPost()`：分流順序 = 指令 → 身分綁定 → AI → 必填攔截 → 寫入 → 回訊；`doGet()` 是瀏覽器狀態頁 | 加新指令、改處理順序 |
| `2.3 喚醒字與補件草稿` | `stripWakeWord()`、`loadDraft/saveDraft/clearDraft`、`mergeDetails()`、`missingRequired()`、`askMissingMessage()` | 改補件話術與 TTL |
| `2.35 身分綁定` | `parseIdentityCommand()`、`bindIdentity()`、`identityLabel()`、`getUserIdentity()`、`fetchLineDisplayName()` | 加新綁定写法、公司/人員邏輯 |
| `2.4 回覆策略` | `decideFlow()`、`shouldReplyFor()`、`isReadOnlyAsking()`、`looksLikeLightData()`、`isReplyAllMode()` | 調整「何時該回訊」 |
| `2.5 / 2.6 指令` | `handleModelCommand()`、`getCandidateModels()`、`handleHelpCommand()` | 改 `/model`、`/help` 文案 |
| `3. Gemini 串接` | `analyzeMessageWithGemini()`：prompt 契約、多模型容錯 | **加欄位時必改 prompt** |
| `4. 核心業務` | `handleDataRouting()` 分流；`doCreate/doUpdate/doRead`；`readData()` 裁寬、`isSummaryRow()`、`matchTemple/matchSpec/normSpec`、`parseDateKey()`、`buildRowValues()`、`renameAgent()`、`setupLightHeader()` | 比對規則與寫入規則 |
| `5. LINE 工具` | `sendLineReply()` | 改訊息格式/加 Flex Message |
| `6./9. 其他` | 狀態頁 HTML、`testGeminiProbe()`、`testLineToken()`；9 區為舊專案殘留 | 排錯用測試筆 |
