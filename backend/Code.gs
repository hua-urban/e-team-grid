/**
 * E 組九宮格展示網站 — 後台 API + 一鍵建表
 * 綁在「E組九宮格_後台資料」這張 Sheet 上。
 *
 * 對照姊妹站 bni-hua-urban 的 Code.gs 風格（setup() / doGet / FIELD 概念）。
 * 規格來源：C:\Second Brain\Obsidian\Joan\工作\BNI-華都更分會\E組九宮格展示網站_PRD.md § 12
 *
 * ★ 第一次部署：選 setup 函式按「執行」→ 自動建表＋灌 15 人初始資料＋授權秘書帳號
 * ★ 前端讀資料：doGet（全員 or ?id=NNN 單人）
 * ★ 前端寫資料：doPost（{id, editor, changes, data}），成功寫回 Sheet + 觸發 Telegram 通知
 *
 * 部署 SOP 見同資料夾 `部署SOP.md`。
 */

// ====== 常數 ======

var SHEET_NAME = 'E組資料';
var HEADERS = ['id', 'name', 'updated_at', 'last_editor', 'data_json'];

// ====== 一鍵初始化 ======

/**
 * setup()
 * 建立新 Spreadsheet「E組九宮格_後台資料」（若尚未綁定既有 Sheet，SpreadsheetApp.create 自建一份）、
 * 建 headers、從 INIT_DATA（15 人，見同檔案尾端常數，data.js 轉換而來）灌初始資料、
 * 加秘書帳號為編輯者、Logger 輸出網址。
 *
 * 執行一次即可；重跑會清空「E組資料」分頁重建（不影響其他分頁）。
 */
function setup() {
  var ss;
  try {
    // 若這支 Code.gs 已經綁在某個 Spreadsheet 容器內（container-bound script），直接用它。
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    ss = null;
  }
  if (!ss) {
    // standalone script（script.new 建立）第一次跑，自建一份新 Sheet。
    ss = SpreadsheetApp.create('E組九宮格_後台資料');
  }

  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();

  var rows = INIT_DATA.map(function (person) {
    return [
      person.id,
      person.name,
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      '（初始匯入）',
      JSON.stringify(person.state)
    ];
  });

  var all = [HEADERS].concat(rows);
  sh.getRange(1, 1, all.length, HEADERS.length).setValues(all);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#C8102E').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.setColumnWidth(5, 500); // data_json 欄加寬方便人工檢視

  // 清掉預設空白分頁（若存在且不是唯一分頁）
  var s1 = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  // 加秘書帳號為編輯者（不影響「誰可以存取 web app」設定，那是部署層另外的鐵律）
  try {
    ss.addEditor('qiongan0208@gmail.com');
  } catch (e) {
    Logger.log('[setup] addEditor 失敗（可能已是編輯者，忽略）：' + e.message);
  }

  Logger.log('[setup] 完成！Sheet URL：' + ss.getUrl());
  Logger.log('[setup] 寫入 ' + rows.length + ' 人（' + INIT_DATA.filter(function (p) {
    return p.state.self.role !== '' || p.state.themes.some(function (t) { return t.service !== ''; });
  }).length + ' 人已建檔內容，其餘為空 81 格待本人填）');

  try {
    SpreadsheetApp.getUi().alert('✅ 已建立/更新「' + SHEET_NAME + '」分頁，共 ' + rows.length + ' 人。\nSheet 網址已印在執行紀錄（Logger）。');
  } catch (e) {
    // 非互動模式（script editor 外呼叫）不彈 alert，靜默通過
  }
}

// ====== 讀取 API：doGet ======

/**
 * doGet(e)
 * 無 ?id= 參數：回全員 { ok:true, data: { "003": {...}, "005": {...}, ... } }
 * 有 ?id=NNN：只回單人 { ok:true, data: {...單人物件...} }（省流量，個人頁用）
 * 找不到分頁 / 找不到該 id → { ok:false, error: "..." }
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    return jsonOutput({ ok: false, error: '找不到分頁：' + SHEET_NAME + '（請先跑 setup()）' });
  }

  var wantId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim() : '';
  var all = readAllRows(sh);

  if (wantId) {
    var person = all[wantId];
    if (!person) {
      return jsonOutput({ ok: false, error: '找不到 id：' + wantId });
    }
    return jsonOutput({ ok: true, data: person });
  }

  return jsonOutput({ ok: true, data: all });
}

/**
 * 讀整張「E組資料」分頁，回傳 { id: { id, name, updated_at, last_editor, ...data_json展開 } }
 */
function readAllRows(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return {};
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idCol = headers.indexOf('id');
  var nameCol = headers.indexOf('name');
  var updatedCol = headers.indexOf('updated_at');
  var editorCol = headers.indexOf('last_editor');
  var dataCol = headers.indexOf('data_json');

  var result = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var id = String(row[idCol] || '').trim();
    if (!id) continue;

    var state = { self: { name: '', en: '', role: '', area: '' }, themes: [] };
    try {
      var parsed = JSON.parse(row[dataCol] || '{}');
      if (parsed && parsed.self) state = parsed;
    } catch (err) {
      Logger.log('[readAllRows] id=' + id + ' data_json 解析失敗，回空 state：' + err.message);
    }

    result[id] = {
      id: id,
      name: row[nameCol] || '',
      updated_at: formatCellDate(row[updatedCol]),
      last_editor: row[editorCol] || '',
      self: state.self,
      themes: state.themes
    };
  }
  return result;
}

/**
 * Sheet 儲存格若已被 Google Sheets 自動辨識成 Date 物件，統一轉字串輸出。
 * （踩坑速查有 GAS Date 陷阱教訓，這裡預先擋，純文字寫入時不會走到這支，但保留防禦）
 */
function formatCellDate(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  }
  return String(v || '');
}

// ====== 寫入 API：doPost ======

/**
 * doPost(e)
 * body（text/plain，前端 fetch 避 CORS preflight）：
 *   { id, editor, changes: [{path,label,oldValue,newValue}], data: { self, themes } }
 * 成功：找到該 id 列 → 覆寫 data_json / updated_at / last_editor → 回 { ok:true }
 *       → 另外觸發 Telegram 通知（失敗不擋寫入，try/catch 分離）
 * 失敗（找不到 id / payload 格式錯）：回 { ok:false, error }
 */
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ ok: false, error: 'payload 不是合法 JSON：' + err.message });
  }

  var id = String(payload.id || '').trim();
  var editor = String(payload.editor || '（未填名）').trim();
  var changes = Array.isArray(payload.changes) ? payload.changes : [];
  var data = payload.data;

  if (!id) return jsonOutput({ ok: false, error: '缺少 id' });
  if (!data || !data.self || !Array.isArray(data.themes)) {
    return jsonOutput({ ok: false, error: 'data 格式不正確（需含 self / themes）' });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return jsonOutput({ ok: false, error: '找不到分頁：' + SHEET_NAME });

  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idCol = headers.indexOf('id');
  var nameCol = headers.indexOf('name');
  var updatedCol = headers.indexOf('updated_at');
  var editorCol = headers.indexOf('last_editor');
  var dataCol = headers.indexOf('data_json');

  var targetRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol] || '').trim() === id) {
      targetRow = i + 1; // sheet 是 1-indexed，values 是 0-indexed
      break;
    }
  }

  if (targetRow === -1) {
    return jsonOutput({ ok: false, error: '找不到 id：' + id + '（Sheet 內無此列，需先跑 setup() 或手動新增）' });
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  var name = data.self.name || values[targetRow - 1][nameCol] || id;

  sh.getRange(targetRow, dataCol + 1).setValue(JSON.stringify(data));
  sh.getRange(targetRow, updatedCol + 1).setValue(now);
  sh.getRange(targetRow, editorCol + 1).setValue(editor);
  // name 欄跟著本人資料同步更新（若本人改了自己的名字）
  if (name) sh.getRange(targetRow, nameCol + 1).setValue(name);

  // Telegram 通知失敗不得阻擋資料寫入已完成的事實，try/catch 分離。
  try {
    notifyTelegram(editor, id, name, changes);
  } catch (notifyErr) {
    Logger.log('[doPost] Telegram 通知失敗（資料已寫入，不影響存檔結果）：' + notifyErr.message);
  }

  return jsonOutput({ ok: true });
}

// ====== Telegram 通知 ======

/**
 * notifyTelegram(editor, id, name, changes)
 * Token / Chat ID 一律從 Script Properties 讀，不寫死在 code 裡（public repo 安全考量）。
 * 部署時設定：專案設定 → 指令碼屬性 → 新增 TG_TOKEN / TG_CHAT_ID
 * 真實值見部署 SOP 指針（見 [[敏感憑證總表]] § Telegram OPS bot）。
 */
function notifyTelegram(editor, id, name, changes) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TG_TOKEN');
  var chatId = props.getProperty('TG_CHAT_ID');

  if (!token || !chatId) {
    Logger.log('[notifyTelegram] TG_TOKEN / TG_CHAT_ID 未設定，略過通知（資料仍正常寫入）');
    return;
  }

  var changeLines = changes.length > 0
    ? changes.map(function (c) {
        var label = c.label || c.path || '（未標示欄位）';
        var newVal = (c.newValue === '' || c.newValue == null) ? '（清空）' : c.newValue;
        return '・' + label + ' → ' + newVal;
      }).join('\n')
    : '（未帶明細）';

  var text = '✏️ E組九宮格更新通知\n'
    + '編輯者：' + editor + '\n'
    + '改了誰的頁：' + name + '（' + id + '）\n'
    + '改動內容：\n' + changeLines;

  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: text }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    Logger.log('[notifyTelegram] Telegram API 回應非 200：' + code + ' ' + resp.getContentText());
  }
}

// ====== 共用工具 ======

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ====== INIT_DATA（setup() 用，見 gen-init-data.js 轉換說明） ======

// INIT_DATA — 由 gen-init-data.js 自動生成，勿手改（改請動 data.js 重跑腳本）
// 來源：C:\Users\Joan\Projects\e-team-grid\data.js（2026-07-06 轉換）
// 15 人：12 人已建檔（含 data.js 全部逐格內容）+ 3 人未建檔（空 81 格，本人可從零填）
var INIT_DATA = [
  {
    "id": "003",
    "name": "劉羿岑",
    "state": {
      "self": {
        "name": "劉羿岑",
        "en": "",
        "role": "節能減碳顧問・ESCO",
        "area": "蒸汽/節能工程"
      },
      "themes": [
        {
          "service": "通風工程",
          "ta": "大空間業主",
          "roles": [
            "水電/機電工廠.營造",
            "消防工程",
            "工安/環安／倉儲物流",
            "室內設計/裝修",
            "塑料品質檢測／健身.食品",
            "高空作業施工／鍚加工.化工",
            "ESCO／百貨.健身.學校",
            "自動控制/IOT／工業.載體"
          ]
        },
        {
          "service": "空調減量",
          "ta": "冷氣商.廠商",
          "roles": [
            "空調工程",
            "消防排煙",
            "建築／備熱",
            "通風工程",
            "蒸氣／IAQ",
            "",
            "ESCO",
            "IOT"
          ]
        },
        {
          "service": "電源監測／節能服務",
          "ta": "製造業",
          "roles": [
            "水電/機電",
            "配電盤工程／資料中心",
            "空調工程／學校.醫院",
            "",
            "ESCO",
            "太陽能/儲能／資料中心",
            "ERP/MES/統／開啟.商辦 ⚠️(?)",
            "ESG/碳管理／百貨/連鎖"
          ]
        },
        {
          "service": "品牌合作",
          "ta": "影音行銷",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "經銷開發",
          "ta": "工班.工程公司",
          "roles": [
            "水電/機電",
            "空調工程",
            "消防排煙",
            "",
            "鐵工/自營 ⚠️(?)",
            "建築/設計/物流／設備商",
            "歐務/工程顧問",
            "ESCO"
          ]
        },
        {
          "service": "永續講座",
          "ta": "學校.政府",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "碳管理服務",
          "ta": "上市櫃.出口廠",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "通風整合服務",
          "ta": "建築師.技師",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        }
      ]
    }
  },
  {
    "id": "005",
    "name": "劉恩秀",
    "state": {
      "self": {
        "name": "劉恩秀",
        "en": "Joan",
        "role": "人力派遣",
        "area": ""
      },
      "themes": [
        {
          "service": "長期派遣",
          "ta": "百貨、品牌",
          "roles": [
            "百貨HR",
            "品牌HR",
            "行銷公司",
            "倉儲HR",
            "品牌業務",
            "房東",
            "活動公關公司HR",
            "客服部HR"
          ]
        },
        {
          "service": "短期派遣",
          "ta": "展場、活動工讀",
          "roles": [
            "百貨",
            "快閃店",
            "行銷公司",
            "工廠產線",
            "品牌業務",
            "場地業務",
            "場地業者",
            "活動公關公司"
          ]
        },
        {
          "service": "平台派遣",
          "ta": "有人力需求的店家",
          "roles": [
            "中小企業",
            "服飾業",
            "家居品牌",
            "醫療器材業者",
            "倉儲",
            "家電業者",
            "房東",
            "律師審合約"
          ]
        },
        {
          "service": "",
          "ta": "想買有5倍報酬都更宅的人",
          "roles": [
            "Ellie老師",
            "仲介",
            "代書",
            "都更土地開發專業人士",
            "BNI夥伴",
            "代銷",
            "置產投資客",
            "建商"
          ]
        },
        {
          "service": "平台上架服務",
          "ta": "有招募需求的企業",
          "roles": [
            "中小企業",
            "銀行",
            "資訊業",
            "裝潢公司",
            "餐飲業",
            "傳播媒體業",
            "電商",
            "保險業"
          ]
        },
        {
          "service": "健康現金流規劃",
          "ta": "銀行、人壽",
          "roles": [
            "銀行",
            "前銀行業主管同事",
            "BNI夥伴",
            "高資產客戶",
            "壽險公司",
            "需要保險服務的人",
            "需要銀行服務的人",
            "中小企業"
          ]
        },
        {
          "service": "企業外包顧問",
          "ta": "需要招募建議的企業",
          "roles": [
            "中小企業",
            "顧問公司",
            "企業HR",
            "客服中心",
            "會計師",
            "物流中心",
            "倉儲工廠",
            "律師"
          ]
        },
        {
          "service": "講座",
          "ta": "求職端／企業端／想了解企業內訓.勞基法的單位",
          "roles": [
            "大專院校",
            "就業服務站",
            "職涯中心",
            "倉儲中心",
            "大型工廠",
            "展覽活動公司",
            "物流中心",
            "客服中心"
          ]
        }
      ]
    }
  },
  {
    "id": "007",
    "name": "張瓊安",
    "state": {
      "self": {
        "name": "張瓊安",
        "en": "Joan",
        "role": "仲介經紀人",
        "area": "三重・蘆洲・新莊"
      },
      "themes": [
        {
          "service": "室內設計｜老宅翻修",
          "ta": "精準控制預算的屋主",
          "roles": [
            "結構技師",
            "拆除工班",
            "水電工班",
            "木作工班",
            "油漆/壁紙",
            "系統櫃廠商",
            "修繕貸款銀行",
            "估價師"
          ]
        },
        {
          "service": "室內設計｜合法隔套",
          "ta": "想快/便宜隔套的二房東",
          "roles": [
            "包租代管公司",
            "租賃仲介",
            "買賣仲介",
            "投資客轉介",
            "水電工班",
            "二房東轉介",
            "木作/隔間工班",
            "泥作工班"
          ]
        },
        {
          "service": "中古屋買賣房產顧問",
          "ta": "雙北買賣屋客戶",
          "roles": [
            "律師",
            "仲介",
            "財稅顧問",
            "代書",
            "銀行房貸專員",
            "土開",
            "估價師",
            "驗屋公司"
          ]
        },
        {
          "service": "自建整合",
          "ta": "100–200坪地主",
          "roles": [
            "都更顧問",
            "結構技師",
            "代書",
            "不動產估價師",
            "建築師",
            "中人",
            "土開",
            "建經公司"
          ]
        },
        {
          "service": "新建案代銷",
          "ta": "30戶以下快銷建商",
          "roles": [
            "買房團/團長",
            "（公設/接待中心）商務空間設計",
            "代銷",
            "仲介",
            "跑單",
            "跑專",
            "建材商",
            "房地產企劃"
          ]
        },
        {
          "service": "租賃代管仲介",
          "ta": "不想自己麻煩的房東",
          "roles": [
            "包租代管公司",
            "清潔公司",
            "維修工班",
            "法律顧問（租賃糾紛）",
            "軟裝電商",
            "家具家電供應商",
            "591/Airbnb 房客來源",
            "物業管理 App"
          ]
        },
        {
          "service": "輕量仲介ERP",
          "ta": "30人以下仲介團隊",
          "roles": [
            "房仲教練培訓講師",
            "仲介同業",
            "工程師",
            "數位廣告投手",
            "AI 相關產業",
            "行銷顧問",
            "SaaS/API 廠商",
            "雲端服務商"
          ]
        },
        {
          "service": "買房收租規劃",
          "ta": "準退休客戶",
          "roles": [
            "保險顧問",
            "退休理財顧問",
            "銀行理專",
            "信託顧問",
            "股票投資人",
            "不動產律師",
            "稅務會計師",
            "仲介同業"
          ]
        }
      ]
    }
  },
  {
    "id": "023",
    "name": "林必涵",
    "state": {
      "self": {
        "name": "林必涵",
        "en": "",
        "role": "廠房資產管理與地主協調",
        "area": "台中農地廠房收租"
      },
      "themes": [
        {
          "service": "廠房出租管理",
          "ta": "持有鐵皮廠房的地主",
          "roles": [
            "廠房管理公司",
            "包租代管業者",
            "廠房仲介",
            "租賃管理顧問",
            "地主社群",
            "",
            "工廠承租戶/二房東",
            "工業地仲介"
          ]
        },
        {
          "service": "地主資產管理",
          "ta": "家族持有農地廠房的二代",
          "roles": [
            "廠登顧問",
            "土地變更顧問",
            "稅務會計師",
            "法人、土地融資銀行",
            "地政士",
            "都市計畫/市地重劃顧問",
            "財務規劃顧問",
            "土地開發"
          ]
        },
        {
          "service": "廠房租約重整",
          "ta": "與二房東交接中的地主",
          "roles": [
            "地政士",
            "不動產律師/租賃律師",
            "跨區域廠房聯盟",
            "",
            "包租代管業者",
            "",
            "",
            "廠房管理公司"
          ]
        },
        {
          "service": "消防改善協調",
          "ta": "有消防疑慮的廠房主",
          "roles": [
            "消防簽證/消防設備",
            "建築師",
            "工安顧問",
            "",
            "使用合法化顧問",
            "",
            "",
            "結構技師"
          ]
        },
        {
          "service": "土地資產評估",
          "ta": "了解自己土地價值及未來發展性的地主",
          "roles": [
            "法人、土地融資銀行",
            "資產傳承顧問",
            "土地投資客",
            "貸款銀行",
            "地方民代/議員資源",
            "商用不動產仲介",
            "資產配置顧問",
            "土地估價師"
          ]
        },
        {
          "service": "廠房接管協調",
          "ta": "管理混亂的出租方",
          "roles": [
            "水電",
            "廠房修繕業者",
            "弱電工程",
            "廠房管理公司",
            "鐵工",
            "跨區域廠房聯盟",
            "電動捲門",
            "監視系統"
          ]
        },
        {
          "service": "廠房活化",
          "ta": "租金效益不佳地主",
          "roles": [
            "二房東管理者",
            "廠辦物管公司",
            "工廠承租戶",
            "倉儲物流業",
            "產業園區業者",
            "租賃管理顧問",
            "在地廠房仲介",
            "地主社群"
          ]
        },
        {
          "service": "廠房風險管理",
          "ta": "擔心租賃糾紛的地主",
          "roles": [
            "不動產律師/租賃律師",
            "火險保險",
            "財產保險",
            "",
            "廠登顧問",
            "消防簽證/消防設備",
            "工安顧問",
            "使用合法顧問"
          ]
        }
      ]
    }
  },
  {
    "id": "025",
    "name": "王佳筠",
    "state": {
      "self": {
        "name": "王佳筠",
        "en": "",
        "role": "建築代銷",
        "area": "北部"
      },
      "themes": [
        {
          "service": "中山區預售案銷售",
          "ta": "有購買能力之買方",
          "roles": [
            "高資產置產2000-4000萬買方",
            "科技廠高階管理層自備資金200-1000萬",
            "馬偕醫院買房團領袖可供介紹費合作",
            "YT或部落客財經/房產/科技/精品",
            "網路廣告商擅長房產、股票、精品等高消費類型",
            "高層KOL短影音曝光/拍攝建案業配",
            "有影響力之平台、召集人",
            "年收入150以上萬以上想買預售案之買方"
          ]
        },
        {
          "service": "新成屋銷售",
          "ta": "建商",
          "roles": [
            "買房團",
            "區域夥伴轉介合作",
            "醫生、律師有大量現金收入的買方",
            "風水老師",
            "高級住宅設計師",
            "地政士會計師",
            "超級業務可轉客源之一",
            "晶華飯店、麗晶精品經理、管理層"
          ]
        },
        {
          "service": "建築案源開發",
          "ta": "即將成案之雙北業主",
          "roles": [
            "雙北市建設公司",
            "Top Sales",
            "地主戶銷售",
            "停車場經營者",
            "建築師",
            "企劃、影片協助提案",
            "地主召集人",
            "公設設計師"
          ]
        },
        {
          "service": "包租代管／代租代管",
          "ta": "大安/大同/中山區屋主",
          "roles": [
            "手上有老屋大安、中山大同",
            "醫師公會",
            "房仲引薦可以置產客",
            "外國人房屋租賃平台",
            "室內設計師",
            "工讀生可以協助帶看",
            "裝潢公司",
            "家具廠商"
          ]
        },
        {
          "service": "新成屋出租管理／Airbnb代管",
          "ta": "",
          "roles": [
            "公司行號需要員工宿舍",
            "醫生工會有限制房屋",
            "日商人資",
            "擅長飯店之清潔人員可以接1-2小時",
            "南港軟體園區人資不 ⚠️(?)",
            "只喜歡買房，不想管理的買方",
            "有Airbnb經驗之屋主交流",
            "信義計劃區外商人資部"
          ]
        },
        {
          "service": "改建投資（大安/大同/中山2000萬內）",
          "ta": "",
          "roles": [
            "Eille老師",
            "青埔專賣宏普系列超級業務",
            "大安、大同中山仲介",
            "北市一線建設公司開發部",
            "高額房貸銀行或房貸單位",
            "都更顧問",
            "土地整合",
            "合資人"
          ]
        },
        {
          "service": "裝修設計師介紹",
          "ta": "",
          "roles": [
            "住宅類設計師",
            "有飯店業經驗的工程",
            "3D建模廠商",
            "系統櫃廠商",
            "家電廠商",
            "油漆廠商",
            "水電工班",
            "軟件廠商"
          ]
        },
        {
          "service": "房屋貸款媒合",
          "ta": "有貸款需求的客戶",
          "roles": [
            "可承做港澳、華僑貸款",
            "套房高成數貸款",
            "高鑑價銀行",
            "增貸資產顧問",
            "高鑑價人壽",
            "資產配置顧問",
            "有推薦費之銀行或人壽",
            "地政士"
          ]
        }
      ]
    }
  },
  {
    "id": "026",
    "name": "楊之慧",
    "state": {
      "self": {
        "name": "楊之慧",
        "en": "",
        "role": "金融業・後台交割人員",
        "area": ""
      },
      "themes": [
        {
          "service": "後台交割",
          "ta": "金融同業",
          "roles": [
            "票券商",
            "債券自營商",
            "券商",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "定期定額-股票ETF諮詢",
          "ta": "",
          "roles": [
            "財經新聞",
            "券商 IPO",
            "公開資訊觀測站",
            "",
            "訂閱財經資訊",
            "",
            "",
            ""
          ]
        },
        {
          "service": "定期定額-個股諮詢",
          "ta": "",
          "roles": [
            "財經新聞",
            "看盤軟體",
            "公開資訊觀測站",
            "",
            "訂閱財經資訊",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "改建投資案媒合",
          "ta": "有潛力物件之改建投資人",
          "roles": [
            "仲介 優質房源",
            "地段 符合改建條件",
            "想推改建的鄰居",
            "",
            "代書/律師（合資）過戶/合約審合",
            "",
            "建商",
            "貸款/保險平台"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "銀行平台引薦",
          "ta": "需信卡.信貸.房貸優質客戶",
          "roles": [
            "國營銀行",
            "民營銀行",
            "壽險業",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "收房租服務",
          "ta": "需租房且有正當職業的客戶",
          "roles": [
            "設計師 輕裝修",
            "有租屋需求的優質客戶",
            "代管出租",
            "",
            "",
            "",
            "",
            ""
          ]
        }
      ]
    }
  },
  {
    "id": "028",
    "name": "林佩君",
    "state": {
      "self": {
        "name": "林佩君",
        "en": "",
        "role": "",
        "area": ""
      },
      "themes": [
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        }
      ]
    }
  },
  {
    "id": "030",
    "name": "李慧如",
    "state": {
      "self": {
        "name": "李慧如",
        "en": "",
        "role": "",
        "area": ""
      },
      "themes": [
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        }
      ]
    }
  },
  {
    "id": "031",
    "name": "凃馨文",
    "state": {
      "self": {
        "name": "凃馨文",
        "en": "Wen",
        "role": "家電 B2B",
        "area": ""
      },
      "themes": [
        {
          "service": "實體通路銷售拓展顧問",
          "ta": "新創消費品牌企業",
          "roles": [
            "電商代營運業者",
            "創業補助專案顧問",
            "商業產品攝影師",
            "會計師",
            "商業影片剪輯師",
            "成熟消費電企業主",
            "商品檢驗公司",
            ""
          ]
        },
        {
          "service": "消費電產品選品開發",
          "ta": "跨領域企業端",
          "roles": [
            "商品檢驗公司",
            "群眾募資平台專案經理",
            "成熟消費電企業主",
            "會計師",
            "商業產品攝影師",
            "",
            "",
            "商業影片剪輯師"
          ]
        },
        {
          "service": "都更潛力物件投資",
          "ta": "台北市公寓持有/投資者",
          "roles": [
            "房仲",
            "都更案投資者",
            "投資區域水電行業者",
            "里長",
            "復康巴士業者",
            "投資區域街邊店經營者",
            "投資區域團媽社團",
            "居家長照機構業者"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "收納整理諮詢",
          "ta": "無痛搬家/斷捨需求者",
          "roles": [
            "房仲",
            "住宅室內設計師",
            "搬家公司",
            "裝潢細清業者",
            "心理諮商師",
            "",
            "",
            "都更改建整合業者"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "小資裝修諮詢",
          "ta": "精打細算成家者",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "（右下，居中格無獨立主題）",
          "ta": "",
          "roles": [
            "房仲",
            "裝修社團",
            "銀行房貸專員",
            "",
            "統包商",
            "",
            "",
            "二手家具業者"
          ]
        }
      ]
    }
  },
  {
    "id": "033",
    "name": "宋馨巧",
    "state": {
      "self": {
        "name": "宋馨巧",
        "en": "",
        "role": "韓語線上教學",
        "area": ""
      },
      "themes": [
        {
          "service": "韓語一對一家教",
          "ta": "任何想學韓語的人",
          "roles": [
            "補習班",
            "語言交換社團",
            "機構",
            "遊學代辦",
            "遠距辦公/線上硬體設備商",
            "外語app廠商",
            "韓語學習 IG 網紅/KOL",
            ""
          ]
        },
        {
          "service": "韓語團體教學",
          "ta": "任何想學韓語的人",
          "roles": [
            "學校",
            "語言交換社團",
            "機構",
            "遊學代辦",
            "第二外語培訓機構",
            "外語app廠商",
            "韓語學習 IG 網紅/KOL",
            ""
          ]
        },
        {
          "service": "線上華語教學",
          "ta": "想學華語的韓國人",
          "roles": [
            "在台韓國留學生社團",
            "遊學代辦",
            "專門招收外籍生之華語中心",
            "華語師資培訓機構",
            "學校",
            "補習班",
            "語言交換社團",
            ""
          ]
        },
        {
          "service": "韓國旅遊規劃",
          "ta": "有韓國旅遊需求的人",
          "roles": [
            "專辦韓國線的旅行社",
            "旅遊保險業務員",
            "韓國自由行攻略 KOL",
            "網路隨身 WiFi 供應商",
            "旅遊周邊品牌商",
            "公司行號",
            "社團/商會",
            ""
          ]
        },
        {
          "service": "韓國醫美服務",
          "ta": "想做醫美的人",
          "roles": [
            "醫美/微整心得分享 KOL",
            "醫美術後保養品牌",
            "韓國在地醫美診所窗口",
            "韓式皮膚管理美容師",
            "時尚/美妝雜誌編輯",
            "想做醫美者",
            "",
            ""
          ]
        },
        {
          "service": "投資改建案",
          "ta": "台北/新北潛力改建物件持有人",
          "roles": [
            "專營雙北中古屋房仲",
            "銀行房貸特約專員",
            "幫租代管",
            "建築師事務所",
            "地政士",
            "室內設計裝潢",
            "水電木工",
            "危老重建/都市更新推動師"
          ]
        },
        {
          "service": "餐飲管理顧問",
          "ta": "需門市管理經驗的店主",
          "roles": [
            "品牌經營者",
            "加盟主",
            "專業連鎖餐飲顧問",
            "房屋仲介",
            "行銷活動企劃公司",
            "餐飲原物料批發商",
            "人力資源相關",
            "食安與法規顧問"
          ]
        },
        {
          "service": "咖啡濾掛包客製",
          "ta": "有送禮需求者",
          "roles": [
            "專業咖啡豆烘焙商",
            "企業福委會/採購窗口",
            "婚禮小物/企劃公司",
            "文創市集主辦方",
            "團購主",
            "商品設計師",
            "",
            ""
          ]
        }
      ]
    }
  },
  {
    "id": "043",
    "name": "洪嘉璿",
    "state": {
      "self": {
        "name": "洪嘉璿",
        "en": "左左",
        "role": "MMT 天賦諮詢師",
        "area": "禾嘉管顧團隊"
      },
      "themes": [
        {
          "service": "社群經營（FB台中租屋社團）",
          "ta": "",
          "roles": [
            "房仲",
            "搬家公司",
            "找店面",
            "包租",
            "一般房東",
            "居家清潔",
            "家電廠商",
            "找廠房的人"
          ]
        },
        {
          "service": "共享空間整合出租",
          "ta": "",
          "roles": [
            "租屋需求的人",
            "樂於分享",
            "愛乾淨的人",
            "需要有室友互動",
            "外地求職者",
            "在意生活品質",
            "",
            ""
          ]
        },
        {
          "service": "MMT天賦諮詢",
          "ta": "",
          "roles": [
            "自媒體引流/用戶體驗設計",
            "業務銷售",
            "個人諮詢/人際關係",
            "團隊主管/企業內訓",
            "一人公司創業者",
            "職涯轉職求職",
            "個人品牌經營者",
            "受眾開發"
          ]
        },
        {
          "service": "財務安全與資產配置規劃",
          "ta": "",
          "roles": [
            "退休金準備討論",
            "車禍調解理賠諮詢",
            "財產保險（車．房）",
            "傳承規劃",
            "投資規劃證券開戶",
            "客戶異業合作",
            "",
            "高額保障與保險資產分析"
          ]
        },
        {
          "service": "水晶礦石工作室／空間出租",
          "ta": "",
          "roles": [
            "療癒課程",
            "礦石產業生態圈（目標）",
            "會議空間",
            "水晶手串",
            "課程．活動",
            "繪畫體驗工作坊似顏繪 ⚠️(?)",
            "3D列印",
            "居家擺飾"
          ]
        },
        {
          "service": "外勞仲介引薦",
          "ta": "",
          "roles": [
            "飯店、行旅業者",
            "工廠廠工",
            "農業農工",
            "長照照顧居家看護",
            "一般企業實習生",
            "家庭幫傭（照顧小孩）",
            "",
            "餐廳需求"
          ]
        },
        {
          "service": "工程與營建風險顧問",
          "ta": "",
          "roles": [
            "室內設計師",
            "土木工程行",
            "水電工程行．師傅",
            "建設公司",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "企業風險管理",
          "ta": "",
          "roles": [
            "外籍人力保險",
            "一般零售店面",
            "公司行號員工團體",
            "工程人員",
            "餐飲",
            "工程行業",
            "甜點品牌",
            "美業"
          ]
        }
      ]
    }
  },
  {
    "id": "047",
    "name": "蔡馨葳",
    "state": {
      "self": {
        "name": "蔡馨葳",
        "en": "Mandy",
        "role": "人壽財務規劃",
        "area": "桃園"
      },
      "themes": [
        {
          "service": "風險管理／財務健檢",
          "ta": "有保險需求的人",
          "roles": [
            "有人事風險規劃需求的人",
            "有財務諮詢／婚前健檢需求的人",
            "醫師（各科）",
            "保險經紀人",
            "護理師",
            "健檢中心業者",
            "長照相關業者／照服員",
            "旅遊業者／旅行社"
          ]
        },
        {
          "service": "團險／職災",
          "ta": "企業主",
          "roles": [
            "公司行號福委會",
            "中小企業主",
            "社區物業經理",
            "企業顧問",
            "社區主委",
            "團康活動業者",
            "團體膳業者",
            "HR"
          ]
        },
        {
          "service": "房產買賣",
          "ta": "手上有二間以上房的人",
          "roles": [
            "仲介經紀",
            "都更整合業",
            "鄰／里長",
            "有 2 間以上房屋所有權人",
            "室內設計師 ⚠️(?)",
            "建設公司",
            "地主／投資客",
            "代書"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "理財課程／教育訓練窗口",
          "ta": "想學財商基礎的人",
          "roles": [
            "教育訓練單位",
            "開課老師",
            "開課空間租借業者",
            "桌遊業者",
            "學務主任（子科系）",
            "想學基礎財商的人",
            "企業顧問",
            "辦活動單位"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "生命靈數教學/諮詢",
          "ta": "想了解個人命理／創業合盤的人",
          "roles": [
            "想了解個人運勢流年規劃／算命的人",
            "網紅／個人 IP",
            "合影做事業的人 ⚠️(?)",
            "業務同業",
            "開課單位",
            "學生職涯發展單位",
            "心理諮商師",
            "爭議活動單位 ⚠️(?)"
          ]
        },
        {
          "service": "稅務規劃",
          "ta": "家庭有分配需求的人",
          "roles": [
            "會計師",
            "律師",
            "家族有分配需求的人",
            "房東（已租代管業）",
            "個人資產超過三千萬的人",
            "地主招集人",
            "銀行貸款人員",
            "銀行信託人員"
          ]
        }
      ]
    }
  },
  {
    "id": "050",
    "name": "謝夢瑰",
    "state": {
      "self": {
        "name": "謝夢瑰",
        "en": "Dian",
        "role": "北市高資產房地產顧問",
        "area": "台北"
      },
      "themes": [
        {
          "service": "房地產買賣仲介",
          "ta": "北市5000萬以上物件買賣方",
          "roles": [
            "老客戶要求/轉介紹",
            "CCIM合作（美國國際認證商業不動產投資師）",
            "老同事合作",
            "貴人引薦 ex合作律師/會計師/室內設計師 etc",
            "扶輪社友",
            "大樓開發經營",
            "同業配合(?) ⚠️(?)",
            "BNI/Zilie 老師"
          ]
        },
        {
          "service": "建地・商辦店面買賣介紹",
          "ta": "全台億以上物件買賣方",
          "roles": [
            "老客戶",
            "建商找地",
            "CCIM 配案",
            "各縣市公會/全聯會合作",
            "扶輪社友",
            "貴人引薦",
            "同業合作",
            "BNI/Zilie 老師"
          ]
        },
        {
          "service": "綜合節稅規劃・農地介紹",
          "ta": "",
          "roles": [
            "老客戶",
            "貴人引薦 VIP ex私人銀行",
            "律師",
            "同業合作",
            "會計師/CFP",
            "銀行理事",
            "地政士（代書）",
            "農地課程同學"
          ]
        },
        {
          "service": "RE/MAX加盟房仲服務",
          "ta": "",
          "roles": [
            "公會會員公司會員代表",
            "房仲店頭拜訪",
            "各品牌超業出來加盟開店",
            "V.S 都更整合",
            "包租代營業+費繳加盟",
            "豪宅租貸公司",
            "外商租貸公司",
            "商務中心"
          ]
        },
        {
          "service": "海外房地產仲介（RE/MAX）",
          "ta": "",
          "roles": [
            "老客戶需求/轉介",
            "移民、黃金簽證、第二家園置產",
            "律師",
            "扶輪社/BNI",
            "會計師",
            "廣告投放",
            "小房蓄學 ⚠️(?)，給留學孩子住或收租",
            "私人、跨國銀行"
          ]
        },
        {
          "service": "陪跑轉職教練／房仲線上課程",
          "ta": "他行超業、想轉戰房仲業者",
          "roles": [
            "招攬事業團隊+經理",
            "線上課程平台",
            "代操/台銷團隊",
            "課後實體諮詢",
            "小編、服務窗口",
            "他行業、超業成功心法拜訪分享",
            "好課聯銷：藍太之課程、Zilie 老師課程",
            "+天賦諮詢師、財富燈塔"
          ]
        },
        {
          "service": "都更宅買賣仲介＋整合協助",
          "ta": "北市都更宅買賣雙方",
          "roles": [
            "爭都更粉絲",
            "主動開發、精準開發",
            "包租代管業者合作",
            "里長",
            "整合業者合作",
            "都更危老建商",
            "代書",
            "同業合作"
          ]
        },
        {
          "service": "建商餘屋銷售",
          "ta": "北市建商",
          "roles": [
            "建商",
            "稽批登記地政士",
            "合建地主戶",
            "貴人引薦",
            "室內設計師",
            "想住一起の買方群",
            "代銷業",
            "廣告投放"
          ]
        }
      ]
    }
  },
  {
    "id": "052",
    "name": "張怡蘋",
    "state": {
      "self": {
        "name": "張怡蘋",
        "en": "April",
        "role": "工程顧問",
        "area": "地主召集人"
      },
      "themes": [
        {
          "service": "包租代管（住家）",
          "ta": "士林區/北投區/松山區的屋主",
          "roles": [
            "包租代管企業聯盟",
            "包租代管加盟店",
            "房仲",
            "法律顧問",
            "士林區/北投區/松山區的屋主親友",
            "網路社團",
            "搬家/清潔",
            "輕裝修/軟裝"
          ]
        },
        {
          "service": "改建投資",
          "ta": "士林區/北投區/松山區房仲",
          "roles": [
            "華都更會員合資",
            "鄰里長",
            "改建案源加速班",
            "Ellie 老師",
            "房貸增加現金流",
            "法律顧問",
            "包租代管",
            "輕裝修/軟裝"
          ]
        },
        {
          "service": "改建協助整合",
          "ta": "開發商/建設公司",
          "roles": [
            "鄰居",
            "鄰里長",
            "改建顧問",
            "Ellie 老師",
            "建築師",
            "開發商/建設公司",
            "地主召集人",
            "全案管理公司"
          ]
        },
        {
          "service": "包租代管（商辦）借址登記/開會空間",
          "ta": "20人以內團隊/公司",
          "roles": [
            "工程顧問",
            "創業社群",
            "士林區/北投區房仲",
            "法律顧問",
            "輕裝修",
            "稅務/會計師",
            "自媒體工作者",
            "消防設備/會議設備"
          ]
        },
        {
          "service": "改建專業評估",
          "ta": "地主召集人",
          "roles": [
            "Ellie 老師",
            "改建顧問",
            "開發商/建設公司",
            "更新會",
            "建築師",
            "地主召集人",
            "地政士/代書",
            "全案管理公司"
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "改建鑑定估價",
          "ta": "開發商/建設公司",
          "roles": [
            "開發商/建設公司",
            "地政士/代書",
            "技師",
            "更新會",
            "建築師",
            "全案管理公司",
            "地主召集人",
            "估價師"
          ]
        }
      ]
    }
  },
  {
    "id": "055",
    "name": "吳念鎧",
    "state": {
      "self": {
        "name": "吳念鎧",
        "en": "",
        "role": "",
        "area": ""
      },
      "themes": [
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        },
        {
          "service": "",
          "ta": "",
          "roles": [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
          ]
        }
      ]
    }
  }
];