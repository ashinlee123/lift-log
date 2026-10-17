/**
 * 健力日誌 ↔ Google Sheet 連結
 * 貼到「擴充功能 → Apps Script」，再「部署 → 新增部署作業 → 網頁應用程式」。
 * 執行身分：我；存取權：所有人。把部署後的網址貼到 app 的「進度 → Google Sheet 紀錄」。
 */
const SETS = '每組紀錄', DAYS = '訓練日', BACKUP = '備份';

// app → sheet: overwrite the two record tabs and store a full backup
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    writeTable_(ss, SETS, data.sets || []);
    writeTable_(ss, DAYS, data.days || []);
    const bk = data.backup || [];
    const sh = sheet_(ss, BACKUP);
    sh.clearContents();
    if (bk.length) sh.getRange(1, 1, bk.length, 1).setNumberFormat('@').setValues(bk);
    return json_({ok: true, at: new Date().toISOString()});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

// sheet → app: ?action=backup returns the last backup for restoring
function doGet(e) {
  if ((e.parameter.action || '') !== 'backup') return json_({ok: true});
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BACKUP);
  if (!sh || sh.getLastRow() < 2) return json_({ok: true, backup: null});
  const text = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])).join('');
  return json_({ok: true, backup: JSON.parse(text)});
}

function writeTable_(ss, name, rows) {
  const sh = sheet_(ss, name);
  sh.clearContents();
  if (!rows.length) return;
  const w = Math.max.apply(null, rows.map(r => r.length));
  const vals = rows.map(r => { const a = r.slice(); while (a.length < w) a.push(''); return a; });
  sh.getRange(1, 1, vals.length, w).setValues(vals);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, w).setFontWeight('bold');
}
function sheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
