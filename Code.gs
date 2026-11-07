// Patient Home Care BD - Backend (Google Apps Script)
// Run setup() once, then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
const TABLES = ['employees', 'clients', 'income', 'expenses', 'salaries', 'settings', 'users', 'invoices'];

function setup() {
  TABLES.forEach(sh);
  const u = sh('users');
  if (u.getLastRow() < 2) { u.getRange(1, 1, 1, 3).setValues([['id', 'user', 'pass']]); u.appendRow(['u1', 'admin', 'admin123']); }
  DriveApp.getFoldersByName('PHC_BD_Files').hasNext() || DriveApp.createFolder('PHC_BD_Files');
}

function doGet() { return ContentService.createTextOutput('PHC BD API OK'); }

function doPost(e) {
  let out;
  try { out = api(JSON.parse(e.postData.contents)); } catch (x) { out = { error: String(x) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sh(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(n) || ss.insertSheet(n);
}

function rows(n) {
  const v = sh(n).getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0];
  return v.slice(1).map(r => { const o = {}; h.forEach((k, i) => o[k] = String(r[i])); return o; });
}

function api(r) {
  const users = rows('users');
  const ok = users.some(u => u.user === r.user && u.pass === r.pass);
  if (!ok) return { error: 'Invalid user ID or password' };
  if (r.action === 'login') return { ok: true };
  if (r.action === 'account') {
    if (!r.newUser || String(r.newPass).length < 4) throw 'ID/password too short';
    const us = sh('users'), v = us.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][1]) === r.user && String(v[i][2]) === r.pass) {
        us.getRange(i + 1, 2, 1, 2).setNumberFormat('@').setValues([[r.newUser, r.newPass]]);
        return { ok: true };
      }
    }
    throw 'User not found';
  }
  if (r.action === 'all') {
    const d = {};
    TABLES.filter(t => t !== 'users').forEach(t => d[t] = rows(t));
    return d;
  }
  if (r.action === 'save') { if (r.sheet === 'users') throw 'Not allowed'; return { row: save(r.sheet, r.row) }; }
  if (r.action === 'del') { del(r.sheet, r.id); return { ok: true }; }
  if (r.action === 'upload') {
    const it = DriveApp.getFoldersByName('PHC_BD_Files');
    const f = it.hasNext() ? it.next() : DriveApp.createFolder('PHC_BD_Files');
    const file = f.createFile(Utilities.newBlob(Utilities.base64Decode(r.data), r.mime, r.name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { id: file.getId(), name: r.name };
  }
  throw 'Unknown action';
}

function save(name, row) {
  if (TABLES.indexOf(name) < 0) throw 'Bad sheet';
  const s = sh(name);
  row = Object.assign({ id: Utilities.getUuid().slice(0, 8) }, row);
  if (!row.id) row.id = Utilities.getUuid().slice(0, 8);
  let h = s.getLastColumn() ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0] : [];
  Object.keys(row).forEach(k => { if (h.indexOf(k) < 0) { h.push(k); s.getRange(1, h.length).setValue(k); } });
  const ic = h.indexOf('id');
  const ids = s.getLastRow() > 1 ? s.getRange(2, ic + 1, s.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  const i = ids.indexOf(row.id);
  const rowNo = i >= 0 ? i + 2 : s.getLastRow() + 1;
  const rng = s.getRange(rowNo, 1, 1, h.length);
  rng.setNumberFormat('@');
  rng.setValues([h.map(k => row[k] === undefined ? '' : row[k])]);
  return row;
}

function del(name, id) {
  if (name === 'users') throw 'Not allowed';
  const s = sh(name);
  const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const ids = s.getRange(2, h.indexOf('id') + 1, Math.max(s.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  const i = ids.indexOf(String(id));
  if (i >= 0) s.deleteRow(i + 2);
}
