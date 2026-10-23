/**
 * Pulse: client intelligence (server side, Google Apps Script).
 *
 * Add this file to the Apps Script project behind the VGCG Client Hub sheet
 * (Extensions > Apps Script), then:
 *   1. Make getIntel and setIntelStatus callable the same way as getComms:
 *      wherever the web app lists the functions it accepts, add these two,
 *      so they get the same sign-in check as everything else.
 *   2. Deploy > Manage deployments > edit the current deployment > New version.
 *      Keep the same deployment so the web app's URL doesn't change.
 *
 * Reads the "Intelligence" tab, which the client intelligence analyst fills.
 * Columns: ID, Client, Date, Found, Title, Summary, Source, URL, Type, Signal, Action, Status
 */
function getIntel() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Intelligence');
  if (!sh || sh.getLastRow() < 2) return { items: [] };
  var v = sh.getDataRange().getDisplayValues();
  var head = v.shift().map(function (x) { return String(x).trim().toLowerCase(); });
  var items = v.filter(function (r) { return r[0]; }).map(function (r) {
    var o = {};
    head.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
  return { items: items };
}

function setIntelStatus(id, status) {
  var ok = ['new', 'reviewed', 'actioned', 'dismissed'];
  if (ok.indexOf(status) < 0) throw new Error('Unknown status');
  var sh = SpreadsheetApp.getActive().getSheetByName('Intelligence');
  if (!sh || sh.getLastRow() < 2) throw new Error('Item not found');
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
  var i = ids.indexOf(String(id));
  if (i < 0) throw new Error('Item not found');
  sh.getRange(i + 2, 12).setValue(status);
  return true;
}
