/**
 * 國科會經費 FAQ + 國外差旅費試算｜Google Apps Script 後端
 * 版本：2026-11
 *
 * 安裝：
 * 1. 將本檔完整取代 Code.gs。
 * 2. 新增 HTML 檔案，名稱必須是 Index，貼入 Index.html。
 * 3. 若由試算表綁定 Apps Script，可直接使用；獨立專案請在「專案設定 > 指令碼屬性」
 *    設定 SPREADSHEET_ID。
 * 4. 執行 setupSheets() 一次並授權。
 * 5. 服務 > 新增服務，啟用「Drive API」（Advanced Google services）。
 * 6. 重新整理試算表，使用「日支標準管理 > 更新主計總處官方日支標準」。
 * 7. 部署為網頁應用程式，執行身分選「我」，存取權限依機構政策設定。
 *
 * 安全設計：
 * - 網站訪客只能讀取資料、執行試算及增加「問題點閱統計」紀錄。
 * - 不提供訪客新增/修改/刪除 QA、流程、網站、承辦人或日支標準的 API。
 * - 所有知識庫與日支標準內容只能由有試算表編輯權限的人員在試算表維護。
 *
 * 日支標準由管理者從主計總處官方公告 PDF 同步；PDF 透過 Drive API 轉成 Google 文件擷取文字。
 * 若官方 PDF 格式變動、無法轉換或解析筆數不足，程式會中止，不覆蓋既有資料。
 * 生活費一律以 USD 計算；訪客端僅能讀取，不提供更新日支標準的公開功能。
 */

const APP = {
  TITLE: '國科會計畫經費報銷查詢網站',
  SHEETS: {
    qa: { name: 'QA資料庫', headers: ['分類','問題','解答','備註','法規依據','關鍵字'] },
    flows: { name: '報銷流程圖', headers: ['分類','流程名稱','圖片連結','步驟說明','注意事項','排序'] },
    websites: { name: '相關網站', headers: ['分類','網站名稱','網址','簡介'] },
    staff: { name: '承辦人員', headers: ['單位','負責業務','姓名','分機','Email','職務代理人','關鍵字'] },
    rates: { name: '國外日支標準', headers: ['適用年度','國家/地區','城市','幣別','每日標準(USD)','生效日','失效日','備註','公告網址','附件網址','同步時間'] },
    cnRates: { name: '大陸港澳日支標準', headers: ['適用年度','國家/地區','城市','幣別','每日標準(USD)','生效日','失效日','備註','公告網址','附件網址','同步時間'] },
    fx: { name: '匯率資料', headers: ['查詢日期','幣別','現金賣出','即期賣出','採用匯率','採用類型','來源網址','抓取時間'] },
    popular: { name: '熱門問題紀錄', headers: ['點閱時間','分類','問題'] }
  },
  BANK_URL: 'https://rate.bot.com.tw',
  BANK_HISTORY_URL: 'https://rate.bot.com.tw/xrt/history?Lang=zh-TW',
  OFFICIAL_RATE_SOURCES: [
    { type: 'foreign', sheetKey: 'rates', pageUrl: 'https://law.dgbas.gov.tw/NewsContent.aspx?id=21777', effectiveDate: '2027-01-01', announcement: '院授主預字第1150103387號函（115/10/22）' },
    { type: 'mainland', sheetKey: 'cnRates', pageUrl: 'https://law.dgbas.gov.tw/NewsContent.aspx?id=21778', effectiveDate: '2027-01-01', announcement: '院授主預字第1150103387號函（115/10/22）' }
  ]
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('日支標準管理')
    .addItem('更新主計總處官方日支標準', 'syncOfficialPerDiemRates')
    .addItem('查看日支標準同步狀態', 'showPerDiemSyncStatus')
    .addToUi();
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return jsonOutput_(routeGet_(e.parameter));
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  // 僅保留唯讀資料相容入口；不接受外部寫入知識庫。
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  if (body.action === 'view') {
    return jsonOutput_(recordQaView(body.category, body.question));
  }
  return jsonOutput_({ success: false, error: '此端點不接受外部新增或修改資料。' });
}

function routeGet_(p) {
  switch (String(p.action || '').toLowerCase()) {
    case 'data': return getInitialData();
    case 'popular': return getPopularQuestions();
    case 'perdiem': return getPerDiemOptions();
    case 'exchange': return getTaiwanBankRate(p.date, p.currency || 'USD');
    default: return { success: false, error: '未知的 action' };
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('找不到試算表。請設定 Script Property：SPREADSHEET_ID。');
}

/** 第一次安裝執行；不會清除或覆寫既有資料。 */
function setupSheets() {
  const ss = getSpreadsheet_();
  Object.keys(APP.SHEETS).forEach(key => {
    const def = APP.SHEETS[key];
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sh.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold').setBackground('#e9eefb').setWrap(true);
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, def.headers.length);
    } else {
      // 既有表只檢查標題，不改寫使用者資料。
      const lastCol = Math.max(sh.getLastColumn(), 1);
      const existing = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v).trim());
      const missing = def.headers.filter(h => existing.indexOf(h) === -1);
      if (missing.length) {
        sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
        sh.getRange(1, lastCol + 1, 1, missing.length).setFontWeight('bold').setBackground('#e9eefb').setWrap(true);
      }
    }
  });
  return '工作表檢查完成。既有資料未被清除；請使用試算表上方「日支標準管理」選單同步官方資料。';
}

function getInitialData() {
  // 主要頁面不等待熱門問題統計，避免大量點閱紀錄拖慢整個網站。
  const ss = getSpreadsheet_();
  const result = { qa: [], flows: [], websites: [], staff: [], popularQuestions: [], loadWarnings: [] };
  [['qa','qa'],['flows','flows'],['websites','websites'],['staff','staff']].forEach(function(pair) {
    try { result[pair[0]] = readSheetObjectsFrom_(ss, pair[1]); }
    catch (err) {
      const message = APP.SHEETS[pair[1]].name + '：' + (err.message || err);
      console.error('網站資料載入失敗 ' + message);
      result.loadWarnings.push(message);
    }
  });
  return result;
}

function readSheetObjectsFrom_(ss, key) {
  const sh = ss.getSheetByName(APP.SHEETS[key].name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getDisplayValues();
  const headers = values.shift().map(v => String(v).trim());
  return values.filter(row => row.some(v => String(v).trim() !== '')).map(row => {
    const obj = {};
    headers.forEach((header, i) => { if (header) obj[header] = row[i] == null ? '' : String(row[i]); });
    return obj;
  });
}

function readSheetObjects_(key) {
  const def = APP.SHEETS[key];
  const sh = getSpreadsheet_().getSheetByName(def.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getDisplayValues();
  const headers = values.shift().map(v => String(v).trim());
  return values.filter(row => row.some(v => String(v).trim() !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] == null ? '' : String(row[i]));
    return obj;
  });
}

function getPopularQuestions() {
  const sh = getSpreadsheet_().getSheetByName(APP.SHEETS.popular.name);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const map = {};
  rows.forEach(r => {
    const category = String(r[1] || '');
    const question = String(r[2] || '');
    if (!question) return;
    const key = category + '\u0001' + question;
    if (!map[key]) map[key] = { category, question, views: 0 };
    map[key].views++;
  });
  return Object.keys(map).map(k => map[k])
    .sort((a, b) => b.views - a.views)
    .slice(0, 5);
}

/** 只允許記錄既有 QA 的點閱，不允許外部建立 FAQ 內容。 */
function recordQaView(category, question) {
  category = String(category || '').trim();
  question = String(question || '').trim();
  if (!category || !question || category.length > 100 || question.length > 500) {
    return { success: false };
  }
  const known = readSheetObjects_('qa').some(r => r['分類'] === category && r['問題'] === question);
  if (!known) return { success: false, error: '問題不存在' };
  const sh = getSpreadsheet_().getSheetByName(APP.SHEETS.popular.name);
  sh.appendRow([new Date(), safeCell_(category), safeCell_(question)]);
  return { success: true };
}

/** 讀取兩份日支標準工作表；金額欄位統一為 USD。 */
function getPerDiemOptions() {
  const rows = readSheetObjects_('rates').concat(readSheetObjects_('cnRates'));
  return rows.filter(r => r['城市'] && Number(r['每日標準(USD)']) > 0).map(r => ({
    year: String(r['適用年度'] || ''),
    country: String(r['國家/地區'] || ''),
    city: String(r['城市'] || ''),
    currency: 'USD',
    rateUSD: Number(String(r['每日標準(USD)']).replace(/,/g, '')),
    startDate: String(r['生效日'] || ''),
    endDate: String(r['失效日'] || ''),
    note: String(r['備註'] || ''),
    sourceUrl: String(r['公告網址'] || ''),
    attachmentUrl: String(r['附件網址'] || '')
  }));
}

/** 唯讀同步狀態，網站端可安全呼叫。 */
function getPerDiemSyncStatus() {
  const props = PropertiesService.getScriptProperties();
  return {
    lastSync: props.getProperty('PERDIEM_LAST_SYNC') || '',
    foreignCount: Number(props.getProperty('PERDIEM_FOREIGN_COUNT') || 0),
    mainlandCount: Number(props.getProperty('PERDIEM_MAINLAND_COUNT') || 0),
    foreignEffective: props.getProperty('PERDIEM_FOREIGN_EFFECTIVE') || '',
    mainlandEffective: props.getProperty('PERDIEM_MAINLAND_EFFECTIVE') || '',
    source: '行政院主計總處主管法規共用系統'
  };
}

/** 管理者手動執行；僅限目前試算表編輯者。 */
function syncOfficialPerDiemRates() {
  assertSpreadsheetEditor_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 先完整下載及解析兩份附件，兩份都成功才寫入，避免只更新一半。
    const prepared = APP.OFFICIAL_RATE_SOURCES.map(source => prepareOfficialRateSource_(source));
    const results = prepared.map(item => writeOfficialRateRows_(item));
    PropertiesService.getScriptProperties().setProperties({
      PERDIEM_LAST_SYNC: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      PERDIEM_FOREIGN_COUNT: String(results[0].count),
      PERDIEM_MAINLAND_COUNT: String(results[1].count),
      PERDIEM_FOREIGN_EFFECTIVE: results[0].effectiveDate,
      PERDIEM_MAINLAND_EFFECTIVE: results[1].effectiveDate
    }, true);
    const message = results.map(r => r.label + '：成功匯入 ' + r.count + ' 筆（生效日 ' + r.effectiveDate + '）').join('\n');
    SpreadsheetApp.getActive().toast(message, '官方日支標準更新完成', 10);
    return { success: true, message: message, results: results };
  } finally { lock.releaseLock(); }
}

function assertSpreadsheetEditor_() {
  const ss = getSpreadsheet_();
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('無法確認目前 Google 帳號。請直接從已登入的 Google 試算表「日支標準管理」選單執行。');
  const file = DriveApp.getFileById(ss.getId());
  const editors = file.getEditors().map(u => String(u.getEmail() || '').toLowerCase());
  const owner = file.getOwner();
  if (owner) editors.push(String(owner.getEmail() || '').toLowerCase());
  if (editors.indexOf(email) < 0) throw new Error('只有此 Google 試算表的編輯者可以更新官方日支標準。');
}

function showPerDiemSyncStatus() {
  const s = getPerDiemSyncStatus();
  const msg = s.lastSync ? '最後同步：' + s.lastSync + '\n國外地區：' + s.foreignCount + ' 筆（' + s.foreignEffective + ' 生效）\n大陸港澳：' + s.mainlandCount + ' 筆（' + s.mainlandEffective + ' 生效）' : '尚未同步官方日支標準。';
  SpreadsheetApp.getUi().alert('官方日支標準同步狀態', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function prepareOfficialRateSource_(source) {
  const response = UrlFetchApp.fetch(source.pageUrl, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (response.getResponseCode() !== 200) throw new Error('無法讀取主計總處公告頁（HTTP ' + response.getResponseCode() + '）：' + source.pageUrl);
  const pdfUrl = findOfficialPdfUrl_(response.getContentText('UTF-8'), source.pageUrl);
  if (!pdfUrl) throw new Error('公告頁中找不到日支數額表 PDF 附件：' + source.pageUrl);
  const pdfResponse = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (pdfResponse.getResponseCode() !== 200) throw new Error('PDF 下載失敗（HTTP ' + pdfResponse.getResponseCode() + '）：' + pdfUrl);
  const blob = pdfResponse.getBlob().setName('PerDiem_' + source.type + '.pdf');
  const bytes = blob.getBytes();
  if (bytes.length < 1000 || bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70) throw new Error('下載內容不是有效 PDF，為避免錯誤更新已停止。');
  const text = convertPdfToText_(blob);
  const parsed = parseOfficialPerDiemText_(text, source.type);
  const minimum = source.type === 'foreign' ? 30 : 20;
  if (parsed.length < minimum) throw new Error('PDF 解析筆數異常（僅 ' + parsed.length + ' 筆，至少需 ' + minimum + ' 筆），未寫入試算表。請檢查官方 PDF 格式或 OCR 結果。');
  return { source: source, pdfUrl: pdfUrl, parsed: parsed };
}

function writeOfficialRateRows_(prepared) {
  const source = prepared.source, parsed = prepared.parsed, pdfUrl = prepared.pdfUrl;
  const sh = getSpreadsheet_().getSheetByName(APP.SHEETS[source.sheetKey].name);
  const headers = APP.SHEETS[source.sheetKey].headers;
  ensureHeaders_(sh, headers);
  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const sourceCol = headerRow.indexOf('公告網址') + 1;
  const lastRow = sh.getLastRow();
  if (lastRow >= 2 && sourceCol > 0) {
    const old = sh.getRange(2, sourceCol, lastRow - 1, 1).getDisplayValues();
    for (let i = old.length - 1; i >= 0; i--) if (old[i][0] === source.pageUrl) sh.deleteRow(i + 2);
  }
  const rocYear = Number(source.effectiveDate.slice(0, 4)) - 1911;
  const now = new Date();
  const rows = parsed.map(r => [rocYear, r.country, r.city, 'USD', r.amount, source.effectiveDate, '', source.announcement, source.pageUrl, pdfUrl, now]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(2, 5, Math.max(sh.getLastRow() - 1, 1), 1).setNumberFormat('0.00');
  return { label: source.type === 'foreign' ? '國外地區' : '大陸／香港／澳門', count: rows.length, effectiveDate: source.effectiveDate, pdfUrl: pdfUrl };
}

function ensureHeaders_(sh, headers) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const current = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v).trim());
  const missing = headers.filter(h => current.indexOf(h) < 0);
  if (missing.length) sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
}

function findOfficialPdfUrl_(html, pageUrl) {
  const anchors = String(html).match(/<a\b[^>]*>[\s\S]*?<\/a>/gi) || [];
  const candidates = [];
  anchors.forEach(a => {
    const hm = a.match(/href\s*=\s*["']([^"']+)["']/i);
    if (!hm) return;
    const label = decodeHtml_(a.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
    const href = decodeHtml_(hm[1]);
    if (!/pdf/i.test(label + ' ' + href) || /對照表|修正對照/i.test(label + ' ' + href)) return;
    if (!/(日支|旅費|大陸地區)/i.test(label + ' ' + href)) return;
    candidates.push(resolveUrl_(pageUrl, href));
  });
  return candidates[0] || '';
}

function resolveUrl_(base, href) {
  if (/^https?:\/\//i.test(href)) return href;
  if (href.indexOf('//') === 0) return 'https:' + href;
  const origin = base.match(/^https?:\/\/[^/]+/)[0];
  if (href.charAt(0) === '/') return origin + href;
  return base.substring(0, base.lastIndexOf('/') + 1) + href;
}

function decodeHtml_(s) {
  return String(s || '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'").replace(/&nbsp;|&#160;/g, ' ').replace(/&#(\d+);/g, (m, n) => String.fromCharCode(Number(n)));
}

/** 需啟用 Advanced Google service：Drive API。暫存 Google 文件會在讀取後刪除。 */
function convertPdfToText_(pdfBlob) {
  let tempId = '';
  try {
    const file = Drive.Files.create({ name: 'TEMP_PERDIEM_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }, pdfBlob, { fields: 'id' });
    tempId = file.id;
    Utilities.sleep(1500);
    return extractConvertedDocText_(DocumentApp.openById(tempId).getBody());
  } finally {
    if (tempId) { try { Drive.Files.remove(tempId); } catch (e) { console.warn('暫存 Google 文件刪除失敗：' + e); } }
  }
}

function extractConvertedDocText_(body) {
  const lines = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r), cells = [];
        for (let c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText().replace(/\s+/g, ' ').trim());
        lines.push(cells.filter(Boolean).join(' '));
      }
    } else if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const t = child.asParagraph().getText().trim();
      if (t) lines.push(t);
    }
  }
  return lines.join('\n');
}

function parseOfficialPerDiemText_(text, type) {
  const lines = String(text || '').replace(/\r/g, '\n').split(/\n+/).map(x => x.replace(/[\t　]+/g, ' ').replace(/\s+/g, ' ').trim()).filter(Boolean);
  const out = [], seen = {};
  lines.forEach(line => {
    if (/^(編號|序號|名稱|城市|國家|日支數額|單位|附註|備註|說明|中華民國|行政院|中央政府)/.test(line)) return;
    const m = line.match(/^(?:(?:\d{1,3})[.、)）]?\s*)?(.+?)\s+(\d{2,4}(?:,\d{3})?(?:\.\d+)?)\s*$/);
    if (!m) return;
    const label = m[1].replace(/^[.、)）\-—–\s]+|[.、)）\-—–\s]+$/g, '').trim();
    const amount = Number(m[2].replace(/,/g, ''));
    if (!label || !Number.isFinite(amount) || amount < 20 || amount > 2000 || /^(合計|總計|小計|備註|附註)/.test(label)) return;
    let country = type === 'foreign' ? '國外' : '中國大陸';
    if (type === 'mainland') {
      if (/香港|Hong\s*Kong/i.test(label)) country = '香港';
      else if (/澳門|Macau/i.test(label)) country = '澳門';
    }
    const city = label.replace(/^\d{1,3}[.、)）]?\s*/, '').trim();
    const key = country + '|' + city;
    if (seen[key]) return;
    seen[key] = true;
    out.push({ country: country, city: city, amount: amount });
  });
  return out;
}


/**
 * 以臺銀歷史牌告頁抓取指定日期匯率。
 * 預設現金賣出；若該幣別無現金賣出，改採即期賣出並清楚標示。
 * 若指定日期為假日/無報價，最多往前回溯 10 個日曆日。
 */
function getTaiwanBankRate(targetDate, currency) {
  targetDate = String(targetDate || '').trim();
  currency = String(currency || 'USD').toUpperCase();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) throw new Error('日期格式需為 YYYY-MM-DD');
  if (!/^[A-Z]{3}$/.test(currency)) throw new Error('幣別代碼不正確');

  const cache = CacheService.getScriptCache();
  const cacheKey = 'BOT_' + targetDate + '_' + currency;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const target = new Date(targetDate + 'T00:00:00');
  for (let back = 0; back <= 10; back++) {
    const d = new Date(target);
    d.setDate(d.getDate() - back);
    const dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
    const url = APP.BANK_URL + '/xrt/quote/' + dateStr + '/' + currency + '?Lang=zh-TW';
    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      if (response.getResponseCode() !== 200) continue;
      const html = response.getContentText('UTF-8');
      const rate = parseBankRateRow_(html, dateStr);
      if (!rate || (!rate.cashSell && !rate.spotSell)) continue;
      const used = rate.cashSell || rate.spotSell;
      const result = {
        success: true,
        requestedDate: targetDate,
        rateDate: dateStr,
        currency: currency,
        cashSell: rate.cashSell || null,
        spotSell: rate.spotSell || null,
        rate: used,
        rateType: rate.cashSell ? '現金賣出' : '即期賣出（該幣別無現金賣出牌價）',
        sourceUrl: url,
        fetchedAt: new Date().toISOString()
      };
      saveFxRecord_(result);
      cache.put(cacheKey, JSON.stringify(result), 21600);
      return result;
    } catch (err) {
      console.warn('臺銀匯率讀取失敗 ' + dateStr + ': ' + err);
    }
  }
  throw new Error('臺灣銀行未能取得該日期前 10 日內的 ' + currency + ' 匯率。請稍後重試，或至「匯率資料」工作表人工登錄並確認。');
}

function parseBankRateRow_(html, dateStr) {
  const dateSlash = dateStr.replace(/-/g, '/');
  const trs = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  for (const tr of trs) {
    const plain = tr.replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;|&#160;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ').trim();
    if (plain.indexOf(dateSlash) === -1 && plain.indexOf(dateStr) === -1) continue;
    const nums = plain.match(/(?:\d+\.\d+|\d+|－|-)/g) || [];
    // 歷史表欄序：日期、幣別、現金買入、現金賣出、即期買入、即期賣出。
    const numeric = nums.slice(1).filter(x => x !== '-' && x !== '－').map(Number);
    // 日期可能被拆成 2026/11/23 的數字，故改以 td 解析欄位。
    const cells = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(td =>
      td.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim()
    );
    const values = cells.map(v => Number(v.replace(/,/g, ''))).filter(v => Number.isFinite(v) && v > 0);
    if (values.length >= 2) {
      // td 數字通常依序為現金買入、現金賣出、即期買入、即期賣出。
      return { cashSell: values.length >= 2 ? values[1] : null, spotSell: values.length >= 4 ? values[3] : null };
    }
    if (numeric.length >= 2) return { cashSell: numeric[1] || null, spotSell: numeric[3] || null };
  }
  return null;
}

function saveFxRecord_(r) {
  const sh = getSpreadsheet_().getSheetByName(APP.SHEETS.fx.name);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === r.rateDate && String(rows[i][1]) === r.currency) {
      sh.getRange(i + 1, 1, 1, 8).setValues([[
        r.rateDate, r.currency, r.cashSell || '', r.spotSell || '',
        r.rate, r.rateType, r.sourceUrl, new Date(r.fetchedAt)
      ]]);
      return;
    }
  }
  sh.appendRow([r.rateDate, r.currency, r.cashSell || '', r.spotSell || '',
    r.rate, r.rateType, r.sourceUrl, new Date(r.fetchedAt)]);
}

/** 防止文字被試算表解讀為公式。 */
function safeCell_(value) {
  const s = String(value == null ? '' : value).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 管理員可手動執行，檢查工作表是否存在及欄位是否完整。 */
function checkSetup() {
  const ss = getSpreadsheet_();
  return Object.keys(APP.SHEETS).map(k => {
    const d = APP.SHEETS[k], sh = ss.getSheetByName(d.name);
    return d.name + '：' + (!sh ? '缺少工作表' :
      '資料列 ' + Math.max(0, sh.getLastRow() - 1) +
      '；欄位 ' + d.headers.join('、'));
  }).join('\n');
}
  
