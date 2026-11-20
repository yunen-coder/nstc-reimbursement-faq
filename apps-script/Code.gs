/** 國科會計畫經費報銷查詢網站 - Google Apps Script API 後端
 * 綁定原本的 Google 試算表後使用。部署為網頁應用程式：執行身分「我」、存取權「任何人」。
 */
const CONFIG = {
  SHEETS: {
    qa: { name: 'QA資料庫', headers: ['分類', '問題', '解答', '備註', '法規依據', '關鍵字'] },
    flows: { name: '報銷流程圖', headers: ['分類', '流程名稱', '圖片連結', '步驟說明', '注意事項', '排序'] },
    websites: { name: '相關網站', headers: ['分類', '網站名稱', '網址', '簡介'] },
    staff: { name: '承辦人員', headers: ['單位', '負責業務', '姓名', '分機', 'Email', '職務代理人', '關鍵字'] }
  }
};

// Cloudflare Worker 會以 ?action=data / popular 呼叫此入口。
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action === 'data') return json_(getInitialData());
  if (action === 'popular') return json_(getPopularQuestions());
  return json_({ error: '請透過網站 API 呼叫。' });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (e && e.parameter && e.parameter.action === 'view') {
      return json_(recordQaView(payload.category, payload.question));
    }
    return json_({ error: '不支援的操作' });
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG.SHEETS).forEach(function(key) {
    const config = CONFIG.SHEETS[key];
    let sheet = ss.getSheetByName(config.name);
    if (!sheet) sheet = ss.insertSheet(config.name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    } else {
      const lastColumn = sheet.getLastColumn();
      const existing = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(h => String(h).trim());
      const missing = config.headers.filter(h => existing.indexOf(h) === -1);
      if (missing.length) sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#eaf0f8');
    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });
  return '4 個工作表已檢查完成。';
}

function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    qa: readSheet_(ss, CONFIG.SHEETS.qa),
    flows: readSheet_(ss, CONFIG.SHEETS.flows),
    websites: readSheet_(ss, CONFIG.SHEETS.websites),
    staff: readSheet_(ss, CONFIG.SHEETS.staff),
    popularQuestions: getPopularQuestions()
  };
}

function readSheet_(ss, config) {
  const sheet = ss.getSheetByName(config.name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const missing = config.headers.filter(h => headers.indexOf(h) === -1);
  if (missing.length) throw new Error('工作表「' + config.name + '」缺少欄位：' + missing.join('、'));
  return values.slice(1).filter(row => row.some(cell => String(cell).trim() !== '')).map(row => {
    const obj = {};
    config.headers.forEach(header => { obj[header] = row[headers.indexOf(header)] || ''; });
    return obj;
  });
}

function getQaKey_(category, question) {
  const text = String(category || '').trim() + '|' + String(question || '').trim();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return 'qa_views_' + bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function recordQaView(category, question) {
  category = String(category || '').trim();
  question = String(question || '').trim();
  if (!category || !question) throw new Error('問題資料不完整。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = readSheet_(ss, CONFIG.SHEETS.qa);
  const exists = rows.some(row => String(row['分類'] || '').trim() === category && String(row['問題'] || '').trim() === question);
  if (!exists) throw new Error('找不到此 QA 問題。');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const key = getQaKey_(category, question);
    const count = Number(props.getProperty(key) || 0) + 1;
    props.setProperty(key, String(count));
    return { success: true, views: count };
  } finally { lock.releaseLock(); }
}

function getPopularQuestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = readSheet_(ss, CONFIG.SHEETS.qa);
  const props = PropertiesService.getScriptProperties();
  return rows.map(row => ({
    category: row['分類'], question: row['問題'], answer: row['解答'],
    views: Number(props.getProperty(getQaKey_(row['分類'], row['問題'])) || 0)
  })).filter(row => row.views > 0).sort((a, b) => b.views - a.views).slice(0, 5);
}
