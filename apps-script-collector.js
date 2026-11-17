// Mission Green collector with email login codes.
// Paste into Google Apps Script (Extensions > Apps Script) of your Google Sheet.
var USERS = 'Users', EVENTS = 'Events';

function clean(v) {
  v = (v === undefined || v === null) ? '' : String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;          // stops spreadsheet formula injection
}
function secret() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hmac(msg) {
  return Utilities.base64Encode(Utilities.computeHmacSha256Signature(msg, secret()));
}
function newCode() {
  var u = Utilities.getUuid().replace(/-/g, ''), A = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', c = '';
  for (var i = 0; i < 8; i++) c += A.charAt(parseInt(u.substr(i * 2, 2), 16) % A.length);
  return c;
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet(ss, name, header) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() == 0) sh.appendRow(header);
  return sh;
}
function users(ss) {
  return sheet(ss, USERS, ['Email','User ID','Nickname','Group','CodeHash','Created','LastRequest','Failed','LockedUntil']);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActiveSpreadsheet(), out;
    if (d.a == 'request') out = requestCode(ss, d);
    else if (d.a == 'login') out = login(ss, d);
    else if (d.a == 'events') out = events(ss, d);
    else out = { ok: false };
    return json(out);
  } catch (err) {
    return json({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
function doGet() { return ContentService.createTextOutput('Mission Green collector is running.'); }

function requestCode(ss, d) {
  var email = String(d.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !d.consent) return { ok: false };
  if (MailApp.getRemainingDailyQuota() < 1) return { ok: false, busy: true };
  var sh = users(ss), data = sh.getDataRange().getValues(), row = -1, now = new Date();
  for (var i = 1; i < data.length; i++) if (data[i][0] == email) { row = i + 1; break; }
  var code = newCode(), h = hmac(email + ':' + code), nick = clean(String(d.nick || '').slice(0, 30));
  if (row < 0) {
    sh.appendRow([email, 'u' + Utilities.getUuid().slice(0, 8), nick, clean(String(d.fam || '').slice(0, 10)), h, now, now, 0, '']);
  } else {
    var last = data[row - 1][6];
    if (last && now - new Date(last) < 120000) return { ok: true };   // max one email per 2 minutes
    sh.getRange(row, 5, 1, 5).setValues([[h, data[row - 1][5], now, 0, '']]);
    if (nick) sh.getRange(row, 3).setValue(nick);
  }
  MailApp.sendEmail({
    to: email,
    name: 'Mission Green',
    subject: 'Your Mission Green login code',
    body: 'Hello,\n\nYour Mission Green login code is:\n\n    ' + code +
          '\n\nOpen the app, go to the Me tab, enter your email and this code, and tap Sign in.\n' +
          'Requesting a new code makes older codes stop working.\n\n' +
          'If you did not request this, you can ignore this email.\n\nThank you for helping our FIRST LEGO League team!'
  });
  return { ok: true };
}

function login(ss, d) {
  var email = String(d.email || '').trim().toLowerCase();
  var code = String(d.code || '').toUpperCase().replace(/\s/g, '');
  var sh = users(ss), data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] != email) continue;
    var row = i + 1, u = data[i];
    if (u[8] && new Date(u[8]) > new Date()) return { ok: false, locked: true };
    if (u[4] && hmac(email + ':' + code) == u[4]) {
      sh.getRange(row, 8, 1, 2).setValues([[0, '']]);
      return { ok: true, uid: u[1], nick: u[2], fam: u[3], token: hmac('t:' + u[1]) };
    }
    var f = (Number(u[7]) || 0) + 1, lockNow = f >= 5;
    sh.getRange(row, 8, 1, 2).setValues([[lockNow ? 0 : f, lockNow ? new Date(Date.now() + 15 * 60000) : '']]);
    return { ok: false, locked: lockNow };
  }
  return { ok: false };
}

function events(ss, d) {
  if (!d.uid || d.token != hmac('t:' + d.uid)) return { ok: false, auth: true };
  var data = users(ss).getDataRange().getValues(), found = false;
  for (var i = 1; i < data.length; i++) if (data[i][1] == d.uid) { found = true; break; }
  if (!found) return { ok: false, auth: true };
  var sh = sheet(ss, EVENTS, ['Received','Timestamp','User ID','Nickname','Group','Type','Habit','Action','Date','Rating','Comment']);
  var list = (d.events || []).slice(0, 200), now = new Date();
  var rows = list.map(function (x) {
    return [now, clean(x.ts), clean(d.uid), clean(x.nick), clean(x.fam), clean(x.type), clean(x.habit),
            clean(x.action), clean(x.date), clean(x.rating), clean(x.comment)];
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { ok: true };
}
