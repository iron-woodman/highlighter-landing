// Waitlist for the highlighter landing: appends each signup to the "Waitlist" sheet.
// Setup: Google Sheet -> Extensions -> Apps Script -> paste this file -> Deploy -> New deployment ->
// Web app, Execute as: Me, Who has access: Anyone -> copy the /exec URL into WAITLIST_ENDPOINT in index.html.

var SHEET = 'Waitlist';
var EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.website) return reply({ ok: true });  // honeypot filled: a bot, pretend success
  var email = String(p.email || '').trim().toLowerCase();
  if (email.length > 254 || !EMAIL.test(email)) return reply({ ok: false, error: 'bad email' });
  var intent = p.intent === 'preorder' ? 'preorder' : 'waitlist';

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet().appendRow([new Date(), cell(email), intent, cell(String(p.src || '').slice(0, 300))]);
  } finally {
    lock.releaseLock();
  }
  return reply({ ok: true });
}

// A browser opening the /exec URL sees this instead of an error page.
function doGet() {
  return reply({ ok: true, service: 'waitlist' });
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['time', 'email', 'intent', 'source']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Text starting with = + - @ would run as a formula in the sheet; a leading quote keeps it plain text.
function cell(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
