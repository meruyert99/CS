/**
 * Topic A Formatives: results collector for GitHub Pages.
 *
 * Paste this whole file into Extensions → Apps Script of a Google Sheet,
 * then Deploy → New deployment → Web app
 *   Execute as: Me      Who has access: Anyone
 * and send the Web app URL to be added to the formatives.
 *
 * Sheets created automatically:
 *   "Formative 1", "Formative 2", "Formative 3"  one row per student, updated live
 *   "Gradebook"                                  name + mark for each formative + average
 * To remove a test or wrong row, just delete it in the sheet.
 */

const SHEETS = { L1: 'Formative 1', L2: 'Formative 2', L3: 'Formative 3' };
const COL = { L1: 2, L2: 3, L3: 4 }; // Gradebook column for each formative
const HEAD = ['Student ID', 'Name', 'Mark /10', 'Correct', 'Done', 'Total', 'Finished', 'Time (min)', 'Level', 'Last update'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheetName = SHEETS[d.formative];
    if (!sheetName || !d.id) return out_({ ok: false });
    const sh = sheet_(sheetName, HEAD);
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    let row = ids.indexOf(String(d.id)) + 2;
    if (row < 2) row = last + 1;
    sh.getRange(row, 1, 1, HEAD.length).setValues([[
      clean_(d.id), clean_(d.name), num_(d.mark), num_(d.correct), num_(d.done), num_(d.total),
      d.finished ? 'yes' : 'no', Math.round(num_(d.timeMs) / 600) / 100, clean_(d.level), new Date()
    ]]);
    gradebook_(clean_(d.name), COL[d.formative], num_(d.mark));
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const sheetName = SHEETS[(e.parameter || {}).formative];
  if (!sheetName) return out_({ rows: [] });
  const sh = sheet_(sheetName, HEAD);
  const n = sh.getLastRow() - 1;
  if (n < 1) return out_({ rows: [] });
  const rows = sh.getRange(2, 1, n, HEAD.length).getValues()
    .filter(r => r[0] !== '' && r[1] !== '')
    .map(r => ({
      id: String(r[0]), name: String(r[1]), mark: +r[2] || 0, correct: +r[3] || 0, done: +r[4] || 0,
      total: +r[5] || 0, finished: r[6] === 'yes', timeMs: Math.round((+r[7] || 0) * 60000)
    }));
  return out_({ rows: rows });
}

function gradebook_(name, col, mark) {
  if (!name || !col) return;
  const sh = sheet_('Gradebook', ['Name', 'Formative 1', 'Formative 2', 'Formative 3', 'Average /10']);
  const last = sh.getLastRow();
  const names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
  let row = names.indexOf(name.trim().toLowerCase()) + 2;
  if (row < 2) {
    row = last + 1;
    sh.getRange(row, 1).setValue(name);
    sh.getRange(row, 5).setFormula('=IF(COUNT(B' + row + ':D' + row + ')=0,"",ROUND(AVERAGE(B' + row + ':D' + row + '),1))');
  }
  sh.getRange(row, col).setValue(mark);
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#2A2150').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function clean_(v) {
  let s = String(v == null ? '' : v).slice(0, 80);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // never let a name become a formula
  return s;
}
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
