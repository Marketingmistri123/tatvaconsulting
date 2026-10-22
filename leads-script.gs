// Tatva Consulting: lead capture backend (Google Apps Script)
// 1. Create a Google Sheet. Open Extensions > Apps Script and paste this file.
// 2. Change PASSWORD below to your own admin password.
// 3. Deploy > New deployment > type: Web app. Execute as: Me. Who has access: Anyone. Copy the Web app URL.
// 4. Paste the URL into index.html (CFG.endpoint) and into the admin page login.
var PASSWORD = 'CHANGE_THIS_PASSWORD';
var HEAD = ['id','at','name','phone','business','industry','problem','monthly_leads','interest','source','page','status'];

function sh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Leads');
  if (!s) { s = ss.insertSheet('Leads'); s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return out_({ ok: true, service: 'tatva-leads' }); }

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad' }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = sh_();
    if (d.action === 'lead') {
      if (d.website) return out_({ ok: true });
      var L = d.lead || {};
      var n = s.getLastRow();
      if (n > 1) {
        var ids = s.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); });
        if (ids.indexOf(String(L.id)) > -1) return out_({ ok: true, dup: true });
      }
      var vals = HEAD.map(function (h) { return String(L[h] == null ? '' : L[h]).slice(0, 300); });
      s.getRange(n + 1, 1, 1, HEAD.length).setNumberFormat('@').setValues([vals]);
      return out_({ ok: true });
    }
    if (d.password !== PASSWORD) return out_({ ok: false, error: 'auth' });
    if (d.action === 'list') {
      var v = s.getDataRange().getValues();
      var rows = v.slice(1).map(function (r) {
        var o = {}; HEAD.forEach(function (h, i) { o[h] = r[i]; }); return o;
      });
      return out_({ ok: true, rows: rows.reverse() });
    }
    if (d.action === 'update') {
      var last = s.getLastRow();
      if (last > 1) {
        var col = s.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < col.length; i++) {
          if (String(col[i][0]) === String(d.id)) {
            s.getRange(i + 2, HEAD.indexOf('status') + 1).setValue(String(d.status).slice(0, 20));
            return out_({ ok: true });
          }
        }
      }
      return out_({ ok: false, error: 'notfound' });
    }
    return out_({ ok: false, error: 'action' });
  } finally { lock.releaseLock(); }
}
