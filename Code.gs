// StarNet ISP — Backend (Google Apps Script)
var SS_ID = '1efk4t2_Fa-jtjEMFaiMjeCUEIDPiUXU7pkuLDNOWUSs';
var SHEET = 'Sheet1';          // মূল ডেটা শিট
var LOG = 'PaymentLog';        // অ্যাপ নিজে তৈরি করবে
var FIRST_COL = 7;             // G = সর্বশেষ মাস (November)
var LAST_COL = 30;             // AD = January 2025
var MON = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('StarNet ISP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() { return SpreadsheetApp.openById(SS_ID); }

// হেডার থেকে মাসের কলাম বের করা ("janu", "Septm", "2026" ইত্যাদি এলোমেলো নাম সামলায়)
function monthCols_(hdr) {
  var now = new Date(), out = [], m = null, y = 0;
  for (var c = FIRST_COL - 1; c < LAST_COL; c++) {
    var i = MON.indexOf(String(hdr[c]).trim().toLowerCase().slice(0, 3));
    if (i < 0) continue;                       // "2026" সেপারেটর বাদ
    if (m === null) { m = i; y = (m <= now.getMonth()) ? now.getFullYear() : now.getFullYear() - 1; }
    else { m--; if (m < 0) { m = 11; y--; } }
    out.push({ col: c + 1, y: y, m: m, key: y + '-' + ('0' + (m + 1)).slice(-2) });
  }
  return out;                                  // নতুন → পুরাতন
}

function parseDate_(v) {
  if (v instanceof Date) return v;
  var p = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!p) return null;
  var y = +p[3]; if (y < 2000 || y > 2100) return null;   // 10/05/0202 টাইপের ভুল বাদ
  return new Date(y, +p[2] - 1, +p[1]);
}

// একবারেই সব ডেটা (ড্যাশবোর্ড + সব ট্যাব)
function getAll(token) {
  var me = auth_(token);
  var ss = ss_(), v = ss.getSheetByName(SHEET).getDataRange().getValues();
  var cols = monthCols_(v[0]), asc = cols.slice().reverse(), last = asc.length - 1, idx = {}, curMoh = '';
  asc.forEach(function (c, k) { idx[c.key] = k; });
  var cust = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], name = String(r[1]).trim();
    if (!name) continue;
    if (isHeader_(r)) { curMoh = name; continue; }   // BW6, BB1 টাইপের মহল্লা-শিরোনাম সারি
    var fee = Number(r[5]) || 0, paid = [], tot = 0, first = -1;
    asc.forEach(function (c, k) {
      var a = Number(r[c.col - 1]) || 0; paid.push(a); tot += a;
      if (a > 0 && first < 0) first = k;
    });
    if (!fee && !tot) continue;                // BW6, BB1 টাইপের এরিয়া-হেডার সারি
    var s = first, jd = parseDate_(r[2]);
    if (jd) {
      var jk = jd.getFullYear() + '-' + ('0' + (jd.getMonth() + 1)).slice(-2);
      s = (jk in idx) ? idx[jk] : (jd < new Date(asc[0].y, asc[0].m, 1) ? 0 : last);
    }
    var active = fee > 0 && s >= 0;
    var due = active ? Math.max(0, fee * (last - s + 1) - tot) : 0;
    cust.push({ row: i + 1, name: name, id: String(r[4] || '').trim(), fee: fee, start: s, moh: curMoh,
                active: active, paid: paid, tot: tot, due: due, dm: fee ? Math.ceil(due / fee) : 0 });
  }
  var monthly = asc.map(function (c, k) {
    var col = 0, bill = 0;
    cust.forEach(function (x) { col += x.paid[k]; if (x.active && x.start <= k) bill += x.fee; });
    return { k: c.key, l: MN[c.m], y: c.y, col: col, bill: bill, due: Math.max(0, bill - col) };
  });
  var act = cust.filter(function (x) { return x.active; });
  var stats = {
    total: act.length,
    fresh: act.filter(function (x) { return x.start === last; }).length,
    collected: monthly[last].col, billed: monthly[last].bill,
    totalDue: act.reduce(function (a, x) { return a + x.due; }, 0),
    dueCount: act.filter(function (x) { return x.due > 0; }).length
  };
  var log = [], all = [], ls = ss.getSheetByName(LOG), tz = ss.getSpreadsheetTimeZone();
  if (ls && ls.getLastRow() > 1) {
    var lv = ls.getRange(2, 1, ls.getLastRow() - 1, 8).getValues();
    for (var q = lv.length - 1; q >= 0 && all.length < 600; q--) {
      var lr = lv[q], dt = new Date(lr[0]);
      all.push({ n: q + 2, d: Utilities.formatDate(dt, tz, 'd MMM'), iso: Utilities.formatDate(dt, tz, 'yyyy-MM-dd'),
        name: lr[2], id: lr[3], amt: lr[4], m: mkey_(lr[5]), u: lr[6], x: lr[7] === 'cancelled' });
    }
    log = all.filter(function (e) { return !e.x; }).slice(0, 30);
  }
  var mohs = [];
  cust.forEach(function (x) { if (x.moh && mohs.indexOf(x.moh) < 0) mohs.push(x.moh); });
  mohs.sort(function (a, b) { return a.localeCompare(b); });
  var months = cols.map(function (c) { return { k: c.key, l: MN[c.m] + ' ' + c.y }; });
  if (me.role === 'owner') return { cust: cust, monthly: monthly, stats: stats, log: log, all: all, months: months, mohs: mohs };
  // স্টাফ: শুধু প্রয়োজনীয় তথ্য, নিজের কালেকশন লগ
  return {
    cust: cust.map(function (x) { return { row: x.row, name: x.name, id: x.id, fee: x.fee, active: x.active, due: x.due, dm: x.dm, moh: x.moh }; }),
    log: log.filter(function (x) { return String(x.u).toLowerCase() === me.name.toLowerCase(); }), months: months, mohs: mohs };
}

function logSheet_(ss) {
  var s = ss.getSheetByName(LOG);
  if (!s) { s = ss.insertSheet(LOG); s.getRange('F:F').setNumberFormat('@'); s.appendRow(['Date', 'Row', 'Name', 'ID', 'Amount', 'Month', 'User', 'Status']); }
  return s;
}

// পেমেন্ট: ওই মাসের ঘরে টাকা যোগ + লগে এন্ট্রি
function addPayment(token, p) {
  var me = auth_(token, ['owner', 'staff']);
  var amt = Number(p.amount);
  if (!(amt > 0)) throw new Error('টাকার পরিমাণ সঠিক নয়');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), sh = ss.getSheetByName(SHEET);
    var cols = monthCols_(sh.getRange(1, 1, 1, LAST_COL).getValues()[0]);
    var c = cols.filter(function (x) { return x.key === p.month; })[0];
    if (!c) throw new Error('এই মাসের কলাম নেই');
    var name = String(p.name || '').trim(), rr = findRow_(sh, p.row, name, String(p.id || '').trim());
    if (!rr) throw new Error('গ্রাহককে খুঁজে পাওয়া যায়নি, পেজ রিফ্রেশ করুন');
    var cell = sh.getRange(rr, c.col);
    cell.setValue((Number(cell.getValue()) || 0) + amt);
    logSheet_(ss).appendRow([new Date(), rr, name, p.id, amt, p.month, me.name]);
  } finally { lock.releaseLock(); }
  return 'ok';
}

// নতুন গ্রাহক: B=নাম, C=জয়েনিং তারিখ, E=আইডি, F=ফি
function addCustomer(token, f) {
  auth_(token, ['owner', 'staff']);
  var name = String(f.name || '').trim(), fee = Number(f.fee), id = String(f.id || '').trim();
  var moh = String(f.moh || '').trim();
  if (!name || !(fee > 0)) throw new Error('নাম ও মাসিক ফি দিন');
  if (!moh) throw new Error('মহল্লার নাম দিন');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = ss_().getSheetByName(SHEET), v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][1]).trim().toLowerCase() === name.toLowerCase()) throw new Error('এই নামে গ্রাহক আগে থেকেই আছে');
      if (id && String(v[i][4]).trim() === id) throw new Error('এই আইডি আগে থেকেই আছে');
    }
    var row = placeRow_(sh, v, moh), d = new Date();
    if (f.date) { var q = f.date.split('-'); d = new Date(+q[0], +q[1] - 1, +q[2]); }
    sh.getRange(row, 2).setValue(name);
    sh.getRange(row, 3).setValue(d).setNumberFormat('dd/MM/yyyy');
    if (id) sh.getRange(row, 5).setValue(id);
    sh.getRange(row, 6).setValue(fee);
  } finally { lock.releaseLock(); }
  return 'ok';
}

// ================= লগিন ও ইউজার =================
var USERS = 'Users';
var INITIAL_OWNER_PASSWORD = 'ChangeMe123';   // setupUsers() চালানোর আগে বদলে নিন

function hash_(salt, p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + p, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function usersSheet_() {
  var ss = ss_(), s = ss.getSheetByName(USERS);
  if (!s) { s = ss.insertSheet(USERS); s.appendRow(['Username', 'Salt', 'Hash', 'Role']); s.setFrozenRows(1); }
  return s;
}
// Apps Script এডিটর থেকে একবার চালান: owner ইউজার তৈরি হবে
function setupUsers() {
  if (usersSheet_().getLastRow() > 1) return 'Users আগেই আছে';
  putUser_('owner', INITIAL_OWNER_PASSWORD, 'owner');
  return 'owner তৈরি হয়েছে';
}
function putUser_(u, p, role) {
  var s = usersSheet_(), v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === u) {
      if (p) { var salt = Utilities.getUuid(); s.getRange(i + 1, 2, 1, 2).setValues([[salt, hash_(salt, p)]]); }
      s.getRange(i + 1, 4).setValue(role); return;
    }
  }
  var sl = Utilities.getUuid();
  s.appendRow([u, sl, hash_(sl, p), role]);
}
function login(u, p) {
  u = String(u || '').trim().toLowerCase(); p = String(p || '');
  var cache = CacheService.getScriptCache(), fk = 'f_' + u, fails = Number(cache.get(fk)) || 0;
  if (fails >= 5) throw new Error('অনেকবার ভুল হয়েছে, ১০ মিনিট পরে চেষ্টা করুন');
  var v = usersSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === u && v[i][3] !== 'off' && hash_(v[i][1], p) === v[i][2]) {
      var t = Utilities.getUuid() + Utilities.getUuid();
      cache.put('s_' + t, JSON.stringify({ name: u }), 21600);   // ৬ ঘণ্টা
      cache.remove(fk);
      loginLog_(u, 'সফল');
      return { token: t, name: u, role: v[i][3] };
    }
  }
  cache.put(fk, String(fails + 1), 600);
  loginLog_(u, 'ভুল পাসওয়ার্ড');
  throw new Error('ইউজারনেম বা পাসওয়ার্ড ভুল');
}
function logout(token) { CacheService.getScriptCache().remove('s_' + token); }
function whoami(token) { return auth_(token); }

// প্রতিটি কলে টোকেন + ভূমিকা যাচাই (ইউজার বন্ধ করলে সাথে সাথে কার্যকর)
function auth_(token, roles) {
  var s = token && CacheService.getScriptCache().get('s_' + token);
  if (!s) throw new Error('AUTH');
  var name = JSON.parse(s).name, v = usersSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === name && v[i][3] !== 'off') {
      var u = { name: name, role: v[i][3] };
      if (roles && roles.indexOf(u.role) < 0) throw new Error('এই কাজের অনুমতি নেই');
      return u;
    }
  }
  throw new Error('AUTH');
}
function listUsers(token) {
  auth_(token, ['owner']);
  var v = usersSheet_().getDataRange().getValues();
  return v.slice(1).map(function (r) { return { u: r[0], role: r[3] }; });
}
function saveUser(token, o) {
  var me = auth_(token, ['owner']);
  var u = String(o.u || '').trim().toLowerCase(), p = String(o.p || ''), role = o.role;
  if (!/^[a-z0-9_.]{3,20}$/.test(u)) throw new Error('ইউজারনেম: ৩-২০ অক্ষর, ইংরেজি ছোট হাতের অক্ষর/সংখ্যা');
  if (['owner', 'staff', 'off'].indexOf(role) < 0) throw new Error('ভূমিকা ভুল');
  if (u === me.name && role !== 'owner') throw new Error('নিজের OWNER ভূমিকা বদলানো যাবে না');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var exists = usersSheet_().getDataRange().getValues().slice(1).some(function (r) { return String(r[0]).toLowerCase() === u; });
    if ((!exists || p) && p.length < 6) throw new Error('পাসওয়ার্ড কমপক্ষে ৬ অক্ষর');
    putUser_(u, p, role);
  } finally { lock.releaseLock(); }
  return 'ok';
}

function mkey_(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  return String(v);
}

// ভুল পেমেন্ট বাতিল (শুধু OWNER): মাসের ঘর থেকে টাকা কমে, লগে "cancelled" লেখা হয়
function cancelPayment(token, n) {
  auth_(token, ['owner']);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = ss_(), ls = ss.getSheetByName(LOG), r = ls.getRange(n, 1, 1, 8).getValues()[0];
    if (!r[2] || r[7] === 'cancelled') throw new Error('এই এন্ট্রি বাতিল করা যাবে না');
    var sh = ss.getSheetByName(SHEET);
    var cols = monthCols_(sh.getRange(1, 1, 1, LAST_COL).getValues()[0]);
    var c = cols.filter(function (z) { return z.key === mkey_(r[5]); })[0];
    if (!c) throw new Error('মাসের কলাম পাওয়া যায়নি');
    var rr = findRow_(sh, r[1], String(r[2]).trim(), String(r[3]).trim());
    if (!rr) throw new Error('গ্রাহককে শিটে খুঁজে পাওয়া যায়নি, শিটে নিজে ঠিক করুন');
    var cell = sh.getRange(rr, c.col);
    cell.setValue(Math.max(0, (Number(cell.getValue()) || 0) - Number(r[4])));
    ls.getRange(n, 8).setValue('cancelled');
  } finally { lock.releaseLock(); }
  return 'ok';
}


// ================= মহল্লা (শিটের শিরোনাম সারি থেকে) =================
// শিরোনাম সারি: শুধু নাম আছে (যেমন BW6), তারিখ/আইডি/ফি/পেমেন্ট কিছু নেই
function isHeader_(r) {
  if (!String(r[1]).trim() || String(r[2]).trim() || String(r[4]).trim() || Number(r[5])) return false;
  for (var c = FIRST_COL - 1; c < LAST_COL; c++) if (String(r[c]).trim() !== '') return false;
  return true;
}

// সারি ঢুকলে নম্বর বদলায়, তাই নাম (ও আইডি) দিয়ে গ্রাহকের সারি খোঁজা হয়
function findRow_(sh, row, name, id) {
  var v = sh.getRange(1, 1, sh.getLastRow(), 5).getValues();
  function ok(i) {
    var e = String(v[i][4]).trim();
    return String(v[i][1]).trim() === name && (!id || !e || e === id || /^R\d+$/.test(id));
  }
  if (row >= 2 && row - 1 < v.length && ok(row - 1)) return row;
  for (var i = 1; i < v.length; i++) if (ok(i)) return i + 1;
  return 0;
}

// নতুন গ্রাহকের জায়গা: নির্বাচিত মহল্লার শেষ সারির পরে (মহল্লা না থাকলে শেষে নতুন শিরোনামসহ)
function placeRow_(sh, v, moh) {
  var i, h = -1, lastCol = sh.getLastColumn(), PF = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  for (i = 1; i < v.length; i++) {
    if (isHeader_(v[i]) && String(v[i][1]).trim().toLowerCase() === moh.toLowerCase()) { h = i; break; }
  }
  var sample = 0;
  for (i = 1; i < v.length; i++) if (String(v[i][1]).trim() && !isHeader_(v[i])) { sample = i; break; }
  if (h < 0) {
    var hr = 0;
    for (i = 1; i < v.length; i++) if (isHeader_(v[i])) { hr = i; break; }
    var nh = v.length + 1;
    if (nh + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), nh + 1 - sh.getMaxRows());
    if (hr) sh.getRange(hr + 1, 1, 1, lastCol).copyTo(sh.getRange(nh, 1, 1, lastCol), PF, false);
    sh.getRange(nh, 2).setValue(moh);
    if (sample) sh.getRange(sample + 1, 1, 1, lastCol).copyTo(sh.getRange(nh + 1, 1, 1, lastCol), PF, false);
    return nh + 1;
  }
  var end = h;
  for (i = h + 1; i < v.length; i++) {
    if (isHeader_(v[i])) break;
    if (String(v[i][1]).trim()) end = i;
  }
  sh.insertRowAfter(end + 1);
  var src = (end === h ? sample : end) + 1;
  if (src > end + 1) src++;
  if (sample || end > h) sh.getRange(src, 1, 1, lastCol).copyTo(sh.getRange(end + 2, 1, 1, lastCol), PF, false);
  return end + 2;
}

// ================= ব্যাকআপ ও নিরাপত্তা =================
var BACKUP_FOLDER = 'StarNet Backup';
var KEEP = 14;   // সর্বশেষ কয়টি ব্যাকআপ রাখা হবে

// Apps Script এডিটর থেকে একবার চালান: শিট প্রোটেক্ট + প্রতিদিন রাত ২টায় ব্যাকআপ + প্রথম ব্যাকআপ
function setupSecurity() {
  protectSheets_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  return 'ব্যাকআপ চালু: ' + backup_();
}

// Users, PaymentLog, LoginLog: শুধু শিটের মালিক এডিট করতে পারবেন
function protectSheets_() {
  var ss = ss_(), me = Session.getEffectiveUser();
  [USERS, LOG, 'LoginLog'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh || sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) return;
    var p = sh.protect().setDescription('StarNet: শুধু মালিক');
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
}

function backup_() {
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER), f = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var tz = ss_().getSpreadsheetTimeZone(), now = new Date();
  DriveApp.getFileById(SS_ID).makeCopy('Wifi Bill Backup ' + Utilities.formatDate(now, tz, 'yyyy-MM-dd HH-mm'), f);
  var files = f.getFiles(), arr = [];
  while (files.hasNext()) arr.push(files.next());
  arr.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (var i = KEEP; i < arr.length; i++) arr[i].setTrashed(true);
  var t = Utilities.formatDate(now, tz, 'd MMM yyyy, hh:mm a');
  PropertiesService.getScriptProperties().setProperty('lastBackup', t);
  return t;
}
function dailyBackup() { backup_(); }
function backupNow(token) { auth_(token, ['owner']); return backup_(); }

function loginLog_(u, status) {
  try {
    var ss = ss_(), s = ss.getSheetByName('LoginLog');
    if (!s) { s = ss.insertSheet('LoginLog'); s.appendRow(['Time', 'User', 'Result']); }
    s.appendRow([new Date(), u, status]);
  } catch (e) {}
}

function securityInfo(token) {
  auth_(token, ['owner']);
  var ss = ss_(), tz = ss.getSpreadsheetTimeZone(), s = ss.getSheetByName('LoginLog'), logins = [];
  if (s && s.getLastRow() > 1) {
    var n = Math.min(10, s.getLastRow() - 1);
    logins = s.getRange(s.getLastRow() - n + 1, 1, n, 3).getValues().reverse().map(function (r) {
      return { t: Utilities.formatDate(new Date(r[0]), tz, 'd MMM, hh:mm a'), u: r[1], s: r[2] };
    });
  }
  return { backup: PropertiesService.getScriptProperties().getProperty('lastBackup') || '', logins: logins };
}

// নিজের পাসওয়ার্ড বদলানো (OWNER ও STAFF দুজনেই)
function changePassword(token, oldP, newP) {
  var me = auth_(token);
  oldP = String(oldP || ''); newP = String(newP || '');
  if (newP.length < 6) throw new Error('নতুন পাসওয়ার্ড কমপক্ষে ৬ অক্ষর');
  var v = usersSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase() === me.name) {
      if (hash_(v[i][1], oldP) !== v[i][2]) throw new Error('বর্তমান পাসওয়ার্ড ভুল');
      putUser_(me.name, newP, v[i][3]);
      return 'ok';
    }
  }
  throw new Error('AUTH');
}
