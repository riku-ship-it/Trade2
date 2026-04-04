// =====================================================================
// 當沖交易日誌系統 — Google Apps Script 後端
// 部署方式：擴充功能 > Apps Script > 部署 > 新增部署
//   類型選「網頁應用程式」，執行身分「我」，存取權「任何人」
// =====================================================================

// ------------------------------------------------------------------
// 主入口：POST 請求
// 前端使用 Content-Type: text/plain 傳送 JSON 字串，
// 可繞過 CORS preflight，GAS 自動回傳 Access-Control-Allow-Origin: *
// ------------------------------------------------------------------
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    let result;
    switch (action) {
      case 'getData':
        result = getData();
        break;
      case 'addTrade':
        result = addTrade(payload);
        break;
      case 'updateNote':
        result = updateNote(payload);
        break;
      case 'updateTrade':
        result = updateTrade(payload);
        break;
      case 'analyzeAI':
        result = analyzeAI(payload);
        break;
      case 'getDailyNotes':
        result = getDailyNotes();
        break;
      case 'saveDailyNote':
        result = saveDailyNote(payload);
        break;
      default:
        result = { status: 'error', message: '未知的 action：' + action };
    }

    return buildResponse(result);

  } catch (err) {
    return buildResponse({ status: 'error', message: err.message });
  }
}

// GET 請求：簡單存活確認（可在瀏覽器直接測試 URL）
function doGet(e) {
  return buildResponse({ status: 'ok', message: '當沖交易日誌 API 運作中' });
}

// ------------------------------------------------------------------
// 建立 JSON 回應
// ------------------------------------------------------------------
function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================================================================
// action: getData
// 讀取「交易紀錄」工作表，回傳所有資料為 JSON 陣列
// ======================================================================
function getData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('交易紀錄');

  if (!sheet) {
    return { status: 'error', message: '找不到「交易紀錄」工作表' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'ok', trades: [] };
  }

  const rows   = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const trades = rows
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      // A欄日期可能是 Date 物件或字串
      let dateStr = '';
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], 'Asia/Taipei', 'yyyy/MM/dd');
      } else {
        dateStr = String(row[0]);
      }

      // H欄寫入時間
      let writtenAt = '';
      if (row[7] instanceof Date) {
        writtenAt = Utilities.formatDate(row[7], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      } else if (row[7]) {
        writtenAt = String(row[7]);
      }

      return {
        date     : dateStr,
        stock    : String(row[1]),
        dir      : String(row[2]),
        entry    : parseFloat(row[3]) || 0,
        exit     : parseFloat(row[4]) || 0,
        lots     : parseInt(row[5])   || 0,
        pnl      : parseInt(row[6])   || 0,
        writtenAt: writtenAt,
        notes    : row[8] ? String(row[8]) : ''
      };
    });

  return { status: 'ok', trades: trades };
}

// ======================================================================
// action: addTrade
// 將一筆交易寫入「交易紀錄」工作表
// 預期 payload 格式：
// {
//   "action": "addTrade",
//   "date"  : "2026/03/22",   // YYYY/MM/DD
//   "stock" : "友達",
//   "dir"   : "空",            // 多 / 空
//   "entry" : 16.58,
//   "exit"  : 15.35,
//   "lots"  : 7,
//   "pnl"   : 8323            // 前端計算後傳入，GAS 不重算
// }
// ======================================================================
function addTrade(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('交易紀錄');

  if (!sheet) {
    return { status: 'error', message: '找不到「交易紀錄」工作表' };
  }

  // 欄位驗證
  const required = ['date', 'stock', 'dir', 'entry', 'exit', 'lots', 'pnl'];
  for (const field of required) {
    if (payload[field] === undefined || payload[field] === null || payload[field] === '') {
      return { status: 'error', message: '缺少欄位：' + field };
    }
  }

  const now       = new Date();
  const writtenAt = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

  sheet.appendRow([
    payload.date,
    payload.stock,
    payload.dir,
    parseFloat(payload.entry),
    parseFloat(payload.exit),
    parseInt(payload.lots),
    parseInt(payload.pnl),
    writtenAt,
    payload.notes || ''
  ]);

  return { status: 'ok', message: '交易已成功寫入' };
}

// ======================================================================
// action: updateNote
// 依 writtenAt 找到對應交易列，更新 I 欄（交易明細）
// 預期 payload：{ "action": "updateNote", "writtenAt": "...", "notes": "..." }
// ======================================================================
function updateNote(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('交易紀錄');

  if (!sheet) {
    return { status: 'error', message: '找不到「交易紀錄」工作表' };
  }
  if (!payload.writtenAt) {
    return { status: 'error', message: '缺少 writtenAt 欄位' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'error', message: '工作表無資料' };
  }

  const hCol = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  for (let i = 0; i < hCol.length; i++) {
    let cellVal = hCol[i][0];
    if (cellVal instanceof Date) {
      cellVal = Utilities.formatDate(cellVal, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    } else {
      cellVal = String(cellVal);
    }
    if (cellVal === payload.writtenAt) {
      sheet.getRange(i + 2, 9).setValue(payload.notes || '');
      return { status: 'ok', message: '明細已更新' };
    }
  }

  return { status: 'error', message: '找不到對應的交易紀錄' };
}

// ======================================================================
// action: updateTrade
// 依 writtenAt 找到對應交易列，更新所有欄位
// ======================================================================
function updateTrade(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('交易紀錄');

  if (!sheet) {
    return { status: 'error', message: '找不到「交易紀錄」工作表' };
  }
  if (!payload.writtenAt) {
    return { status: 'error', message: '缺少 writtenAt 欄位' };
  }

  const required = ['date', 'stock', 'dir', 'entry', 'exit', 'lots', 'pnl'];
  for (const field of required) {
    if (payload[field] === undefined || payload[field] === null || payload[field] === '') {
      return { status: 'error', message: '缺少欄位：' + field };
    }
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'error', message: '工作表無資料' };
  }

  const hCol = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  for (let i = 0; i < hCol.length; i++) {
    let cellVal = hCol[i][0];
    if (cellVal instanceof Date) {
      cellVal = Utilities.formatDate(cellVal, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    } else {
      cellVal = String(cellVal);
    }
    if (cellVal === payload.writtenAt) {
      const row = i + 2;
      sheet.getRange(row, 1).setValue(payload.date);
      sheet.getRange(row, 2).setValue(payload.stock);
      sheet.getRange(row, 3).setValue(payload.dir);
      sheet.getRange(row, 4).setValue(parseFloat(payload.entry));
      sheet.getRange(row, 5).setValue(parseFloat(payload.exit));
      sheet.getRange(row, 6).setValue(parseInt(payload.lots));
      sheet.getRange(row, 7).setValue(parseInt(payload.pnl));
      sheet.getRange(row, 9).setValue(payload.notes || '');
      return { status: 'ok', message: '交易已更新' };
    }
  }

  return { status: 'error', message: '找不到對應的交易紀錄' };
}

// ======================================================================
// action: analyzeAI（預留）
// 目前回傳 coming_soon；日後呼叫 callGeminiAPI()
// ======================================================================
function analyzeAI(payload) {
  // ── 未來實作流程 ──────────────────────────────────────────────────
  // 1. 從 payload 取得 startDate、endDate
  // 2. 呼叫 getData() 取得所有交易，篩選出該區間資料
  // 3. 計算區間統計（總損益、勝率、各股表現…）
  // 4. 組裝 Prompt 字串
  // 5. result = callGeminiAPI(prompt)
  // 6. 將分析結果寫入「報告存檔」工作表
  // 7. 回傳 { status: 'ok', report: result }
  // ─────────────────────────────────────────────────────────────────
  return { status: 'coming_soon' };
}

// ======================================================================
// Gemini API 預留接口
// API Key 與 Model 透過 PropertiesService 讀取，永不寫死在程式碼內。
//
// 設定方式：
//   GAS 後台 > 專案設定 > 指令碼屬性 > 新增屬性
//   GEMINI_API_KEY  →  你的 Gemini API Key
//   GEMINI_MODEL    →  例如 gemini-1.5-flash
// ======================================================================
function callGeminiAPI(prompt) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const MODEL   = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');

  // 未來在此填入 API 呼叫邏輯，範例架構：
  //
  // const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${API_KEY}`;
  // const options = {
  //   method     : 'post',
  //   contentType: 'application/json',
  //   payload    : JSON.stringify({
  //     contents: [{ parts: [{ text: prompt }] }]
  //   })
  // };
  // const response = UrlFetchApp.fetch(url, options);
  // const json     = JSON.parse(response.getContentText());
  // return json.candidates[0].content.parts[0].text;

  return null;
}

// ======================================================================
// 工具函式：將 AI 分析結果寫入「報告存檔」工作表（供 analyzeAI 未來呼叫）
// ======================================================================
function saveReport(dateRange, reportContent) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('報告存檔');
  if (!sheet) return;

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([dateRange, reportContent, now]);
}

// ======================================================================
// action: getDailyNotes
// 讀取「每日心得」工作表，回傳所有心得為 JSON 陣列
// ======================================================================
function getDailyNotes() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName('每日心得');

  if (!sheet) {
    sheet = ss.insertSheet('每日心得');
    sheet.appendRow(['日期', '心得', '更新時間']);
    return { status: 'ok', notes: [] };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: 'ok', notes: [] };
  }

  const rows  = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const notes = rows
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      let dateStr = '';
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], 'Asia/Taipei', 'yyyy/MM/dd');
      } else {
        dateStr = String(row[0]);
      }
      let updatedAt = '';
      if (row[2] instanceof Date) {
        updatedAt = Utilities.formatDate(row[2], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      } else if (row[2]) {
        updatedAt = String(row[2]);
      }
      return { date: dateStr, note: String(row[1] || ''), updatedAt };
    });

  return { status: 'ok', notes };
}

// ======================================================================
// action: saveDailyNote
// 新增或更新「每日心得」工作表中某日的心得
// 預期 payload：{ "action": "saveDailyNote", "date": "2026/04/04", "note": "..." }
// ======================================================================
function saveDailyNote(payload) {
  if (!payload.date) {
    return { status: 'error', message: '缺少 date 欄位' };
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName('每日心得');
  if (!sheet) {
    sheet = ss.insertSheet('每日心得');
    sheet.appendRow(['日期', '心得', '更新時間']);
  }

  const now       = new Date();
  const updatedAt = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const noteText  = payload.note || '';
  const dateKey   = payload.date;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const aCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < aCol.length; i++) {
      let cellVal = aCol[i][0];
      if (cellVal instanceof Date) {
        cellVal = Utilities.formatDate(cellVal, 'Asia/Taipei', 'yyyy/MM/dd');
      } else {
        cellVal = String(cellVal);
      }
      if (cellVal === dateKey) {
        sheet.getRange(i + 2, 2).setValue(noteText);
        sheet.getRange(i + 2, 3).setValue(updatedAt);
        return { status: 'ok', message: '心得已更新' };
      }
    }
  }

  sheet.appendRow([dateKey, noteText, updatedAt]);
  return { status: 'ok', message: '心得已儲存' };
}
