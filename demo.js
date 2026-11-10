/* Giả lập Google Apps Script (SpreadsheetApp, LockService...) để chạy Code.gs trong trình duyệt/Node.
   Chỉ dùng cho chế độ DEMO và test. Dữ liệu lưu ở localStorage của trình duyệt. */
var __LS = (function () { try { return window.localStorage; } catch (e) { return null; } })();
var __DB = (function () { try { return JSON.parse(__LS.getItem('eoq_demo_db5')) || {}; } catch (e) { return {}; } })();
var DEMO_NOW = (function () { try { return __LS.getItem('eoq_demo_now') || ''; } catch (e) { return ''; } })();
function __save() { try { __LS.setItem('eoq_demo_db5', JSON.stringify(__DB)); } catch (e) {} }
function __cell(v) { return (typeof v === 'string' && v.charAt(0) === "'") ? v.slice(1) : v; }
function __Sheet(name) { this.name = name; }
__Sheet.prototype.rows = function () { return __DB[this.name]; };
__Sheet.prototype.getLastRow = function () { return this.rows().length; };
__Sheet.prototype.appendRow = function (r) { this.rows().push(r.map(__cell)); };
__Sheet.prototype.setFrozenRows = function () {};
__Sheet.prototype.clearContents = function () { this.rows().length = 0; };
__Sheet.prototype.getLastColumn = function () { return (this.rows()[0] || []).length; };
__Sheet.prototype.getDataRange = function () {
  var rows = this.rows();
  return { getValues: function () { return JSON.parse(JSON.stringify(rows)); } };
};
__Sheet.prototype.getRange = function (row, col, nr, nc) {
  var rows = this.rows();
  return {
    getValues: function () {
      var out = [];
      for (var i = 0; i < nr; i++) { var r = rows[row - 1 + i] || []; out.push(r.slice(col - 1, col - 1 + nc)); }
      return JSON.parse(JSON.stringify(out));
    },
    setValues: function (vals) {
      for (var i = 0; i < nr; i++) { var r = rows[row - 1 + i] || (rows[row - 1 + i] = []); for (var j = 0; j < nc; j++) r[col - 1 + j] = __cell(vals[i][j]); }
    }
  };
};
var __SS = {
  getSheetByName: function (n) { return __DB[n] ? new __Sheet(n) : null; },
  insertSheet: function (n) { __DB[n] = []; return new __Sheet(n); }
};
var SpreadsheetApp = { getActiveSpreadsheet: function () { return __SS; }, getUi: function () { return { alert: function () {} }; } };
var LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
var ContentService = { MimeType: { JSON: 'json' }, createTextOutput: function (s) { return { content: s, setMimeType: function () { return this; } }; } };
var Utilities = { getUuid: function () { return 'xxxxxxxxxxxxxxxxxxxxxxxx'.replace(/x/g, function () { return (Math.random() * 16 | 0).toString(16); }); } };
var CacheService = { getScriptCache: function () { return { get: function () { return null; }, put: function () {}, remove: function () {} }; } };

var __MAILS = [];
var MailApp = { sendEmail: function (o) { __MAILS.push(o); } };
/**
 * Adtima – Employee of the Quarter · Backend (Google Apps Script) · v2
 * Dữ liệu nằm trong Google Sheet gắn với script này.
 * Chạy setup() một lần để tạo các tab. Deploy: Web app · Execute as Me · Anyone.
 */

/** Bản TEST: đặt true để đề cử và chấm điểm luôn mở (bỏ qua mốc thời gian trong Settings). Bản thật để false. */
var TEST_MODE = false;

var SHEETS = {
  Settings: ['key', 'value', 'note'],
  Employees: ['domain', 'name', 'title', 'unit', 'role', 'accessCode', 'active'],
  Nominations: ['id', 'createdAt', 'nominatorDomain', 'nominatorName', 'unit', 'candidateDomain', 'candidateName',
    'candidateTitle', 'sow', 'ev1', 'ev2', 'ev3', 'ev4', 'r1', 'r2', 'r3', 'r4', 'status', 'hrNote', 'updatedAt'],
  Votes: ['id', 'createdAt', 'voterDomain', 'voterName', 'voterUnit', 'nominationId', 'candidateDomain', 'candidateName',
    'candidateUnit', 'c1', 'c2', 'c3', 'c4', 'total', 'comment']
};

var DEFAULT_SETTINGS = [
  ['cycleName', 'Q3/2026', 'Tên đợt hiển thị trên web'],
  ['nominationStart', '2026-11-10 09:00', 'Giờ Việt Nam, định dạng YYYY-MM-DD HH:mm'],
  ['nominationEnd', '2026-11-12 23:59', 'Đóng đề cử hết ngày 12/11'],
  ['votingStart', '2026-11-13 00:00', 'Mở chấm điểm'],
  ['votingEnd', '2026-11-17 23:59', 'Đóng chấm điểm hết ngày 17/11'],
  ['finalizeDate', '18/11', 'Chỉ để hiển thị'],
  ['ceremonyDate', '19/11', 'Chỉ để hiển thị'],
  ['prize', '2.000.000 VNĐ', ''],
  ['minJudges', '2', 'Số giám khảo tối thiểu để được xếp hạng'],
  ['rankingMethod', 'average', 'average = điểm trung bình các phiếu; sum = tổng điểm']
];


var CRITERIA = [
  { key: 'c1', ev: 'ev1', r: 'r1', name: 'Performance & Delivery', weight: 35 },
  { key: 'c2', ev: 'ev2', r: 'r2', name: 'Impact & Contribution', weight: 30 },
  { key: 'c3', ev: 'ev3', r: 'r3', name: 'Ownership & Proactiveness', weight: 20 },
  { key: 'c4', ev: 'ev4', r: 'r4', name: 'Improvement & Innovation', weight: 15 }
];

/** Người bị chặn chấm cả nhóm team (khớp theo đầu tên team). Cú pháp: domain=Tiền tố1,Tiền tố2 ; domain2=...
 *  Mặc định: không chặn ai. Muốn chặn, thêm dòng voteBlock vào tab Settings, ví dụ: vinhnq5=Sales */
var VOTE_BLOCK_DEFAULT = '';
function voteBlock_(domain) {
  var raw = settings_().voteBlock;
  raw = (raw === undefined || str_(raw) === '') ? VOTE_BLOCK_DEFAULT : str_(raw);
  var out = [];
  raw.split(';').forEach(function (part) {
    var kv = part.split('=');
    if (kv.length < 2 || dom_(kv[0]) !== dom_(domain)) return;
    kv[1].split(',').forEach(function (x) { x = x.trim().toLowerCase(); if (x && x !== 'none') out.push(x); });
  });
  return out;
}

var LIMITS = { sow: 1000, ev: 1000, comment: 1000, note: 500 };

/* ---------------- Setup, menu, auto mã ---------------- */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[name]);
      sh.setFrozenRows(1);
    }
  });
  var st = ss.getSheetByName('Settings');
  if (st.getLastRow() === 1) DEFAULT_SETTINGS.forEach(function (r) { st.appendRow(r); });
  return 'OK';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('EOQ')
    .addItem('Tạo bảng tính (setup)', 'setup')
    .addItem('Tạo mã đăng nhập còn thiếu', 'generateCodesMenu')
    .addToUi();
}

/** Tự tạo mã khi HR dán/sửa danh sách ở tab Employees. */
function onEdit(e) {
  try {
    var nm = e && e.range && e.range.getSheet().getName();
    if (nm) cacheDrop_(nm);
    if (nm === 'Employees') generateCodes_();
  } catch (err) { /* bỏ qua: vẫn có menu EOQ để tạo mã thủ công */ }
}

function generateCodesMenu() {
  var n = generateCodes_();
  SpreadsheetApp.getUi().alert('Đã tạo ' + n + ' mã đăng nhập mới (cột accessCode, tab Employees).');
}

/* ---------------- HTTP ---------------- */

function doGet() {
  return json_({ ok: true, data: { service: 'Adtima EOQ' } });
}

function doPost(e) {
  var res;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = { ok: true, data: route_(req) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return json_(res);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  var a = req.action;
  EMP_CACHE_ = null;
  if (a === 'requestCode') return requestCode_(req.domain);
  var me = auth_(req);
  INCLUDE_SELF_ = !!req.includeSelf;
  switch (a) {
    case 'login': return me.role === 'leader' ? leaderAll_(me) : home_(me);
    case 'leaderAll': return leaderAll_(me);
    case 'leaderHome': return home_(me);
    case 'candidates': return publicData_();
    case 'submitNomination': return withLock_(function () { return submitNomination_(me, req.data || {}); });
    case 'withdrawNomination': return withLock_(function () { return withdrawNomination_(me, req.id); });
    case 'ballot': return ballot_(me);
    case 'submitVote': return withLock_(function () { return submitVote_(me, req.entries || []); });
    case 'adminData': requireAdmin_(me); return adminData_();
    case 'adminReview': requireAdmin_(me); return withLock_(function () { return adminReview_(req.id, req.status, req.note); });
    case 'adminImport': requireAdmin_(me); return withLock_(function () { return adminImport_(req.rows || [], req.mode); });
    case 'adminUpdateEmployee': requireAdmin_(me); return withLock_(function () { return adminUpdateEmployee_(me, req.domainKey, req.data || {}); });
    case 'adminDelete': requireAdmin_(me); return withLock_(function () { return adminDelete_(req.id); });
    case 'adminSendCodes': requireAdmin_(me); return sendCodes_(req.domains);
    case 'adminEdit': requireAdmin_(me); return withLock_(function () { return adminEdit_(req.id, req.data || {}); });
  }
  throw new Error('Thao tác không hợp lệ.');
}

/* ---------------- Gửi mã qua email ---------------- */

/** Email = domain@emailDomain. Đổi đuôi email bằng dòng emailDomain trong tab Settings. Thêm dòng webUrl để chèn link web vào email. */
var MAIL_DOMAIN_DEFAULT = 'vng.com.vn';

function mailOf_(u) {
  var s = settings_();
  return u.domain + '@' + (str_(s.emailDomain) || MAIL_DOMAIN_DEFAULT);
}

function sendCodeMail_(u) {
  var s = settings_();
  var url = str_(s.webUrl);
  var cycle = str_(s.cycleName) || 'Employee of the Quarter';
  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#14213D;line-height:1.6">' +
    '<p>Chào ' + esc_(u.name) + ',</p>' +
    '<p>Đây là thông tin đăng nhập nền tảng đề cử &amp; bình chọn <b>Employee of the Quarter ' + esc_(cycle) + '</b> của Adtima:</p>' +
    '<p style="background:#F2F7FF;border-radius:8px;padding:12px 16px">Domain: <b>' + esc_(u.domain) + '</b><br>Mã đăng nhập: <b style="font-size:18px;letter-spacing:2px">' + esc_(u.accessCode) + '</b></p>' +
    (url ? '<p>Truy cập: <a href="' + esc_(url) + '">' + esc_(url) + '</a></p>' : '') +
    '<p>Mã này dành riêng cho bạn, vui lòng không chia sẻ. Cần hỗ trợ, liên hệ HRBP.</p></div>';
  MailApp.sendEmail({ to: mailOf_(u), subject: '[Adtima] Mã đăng nhập Employee of the Quarter ' + cycle, htmlBody: html, name: 'Adtima HR' });
}

function esc_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

/** Leader tự nhận lại mã: không cần đăng nhập. Luôn trả cùng một thông báo để không lộ domain nào tồn tại; tối đa 1 email / 5 phút / domain. */
function requestCode_(domain) {
  var d = dom_(domain);
  if (!d) throw new Error('Vui lòng nhập domain.');
  var msg = { message: 'Nếu domain hợp lệ, mã đăng nhập đã được gửi về email ' + d + '@' + (str_(settings_().emailDomain) || MAIL_DOMAIN_DEFAULT) + '. Hãy kiểm tra cả mục Spam.' };
  var u = employees_()[d];
  if (!u || !canLogin_(u.role) || !u.active || !u.accessCode) return msg;
  var cache = null;
  try { cache = CacheService.getScriptCache(); if (cache.get('rc_' + d)) return msg; cache.put('rc_' + d, '1', 300); } catch (e) {}
  try { sendCodeMail_(u); } catch (e) { throw new Error('Chưa gửi được email. Liên hệ HRBP.'); }
  return msg;
}

/** Admin gửi mã: domains rỗng = tất cả Team Leader / Sub-team Leader đang hoạt động. */
function sendCodes_(domains) {
  var emp = employees_();
  var want = Array.isArray(domains) && domains.length ? domains.map(dom_) : null;
  var sent = 0, failed = [];
  Object.keys(emp).forEach(function (k) {
    var u = emp[k];
    if (!u.active || !u.accessCode || !canLogin_(u.role)) return;
    if (want ? want.indexOf(u.domain) < 0 : u.role === 'admin') return;
    try { sendCodeMail_(u); sent++; } catch (e) { failed.push(u.domain); }
  });
  return { sent: sent, failed: failed };
}

/** Chạy hàm này 1 lần trong trình soạn thảo Apps Script (chọn authorizeMail → Run) để cấp quyền gửi email và gửi thử 1 mail cho chính bạn. */
function authorizeMail() {
  var me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to: me, subject: '[Adtima EOQ] Thử gửi email', htmlBody: 'Quyền gửi email đã hoạt động. Hạn mức còn lại hôm nay: ' + MailApp.getRemainingDailyQuota() + ' email.', name: 'Adtima HR' });
  Logger.log('Đã gửi email thử tới ' + me);
}

/** Phiếu còn hiệu lực: bỏ phiếu của đề cử đã bị Admin xoá (phiếu vẫn nằm trong Sheet để lưu vết). */
function liveVotes_(noms) {
  var gone = {};
  (noms || rows_('Nominations')).forEach(function (n) { if (n.status === 'deleted') gone[str_(n.id)] = 1; });
  return rows_('Votes').filter(function (v) { return !gone[str_(v.nominationId)]; });
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- Sheet helpers ---------------- */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Thiếu tab "' + name + '". Hãy chạy setup().');
  return sh;
}

var CACHED_SHEETS_ = { Employees: 1, Settings: 1 };
function cacheGet_(name) {
  try { var v = CacheService.getScriptCache().get('rows_' + name); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function cachePut_(name, values) {
  try { var j = JSON.stringify(values); if (j.length < 95000) CacheService.getScriptCache().put('rows_' + name, j, 60); } catch (e) {}
}
function cacheDrop_(name) {
  try { CacheService.getScriptCache().remove('rows_' + name); } catch (e) {}
}

function rows_(name) {
  var values = CACHED_SHEETS_[name] ? cacheGet_(name) : null;
  if (!values) {
    values = sheet_(name).getDataRange().getValues().map(function (r) { return r.map(function (c) { return c instanceof Date ? c.toISOString() : c; }); });
    if (CACHED_SHEETS_[name]) cachePut_(name, values);
  }
  values = values.slice();
  var head = (values.shift() || []).map(function (h) { return String(h).trim(); });
  return values.map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  }).filter(function (o) { return head.some(function (h) { return o[h] !== '' && o[h] !== null; }); });
}

function safe_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function append_(name, obj) {
  cacheDrop_(name);
  sheet_(name).appendRow(SHEETS[name].map(function (h) { return obj[h] === undefined ? '' : safe_(obj[h]); }));
}

function update_(name, row, obj) {
  cacheDrop_(name);
  if (name === 'Employees') EMP_CACHE_ = null;
  var sh = sheet_(name);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn ? sh.getLastColumn() : SHEETS[name].length).getValues()[0].map(function (h) { return String(h).trim(); });
  var cur = sh.getRange(row, 1, 1, head.length).getValues()[0];
  head.forEach(function (h, j) { if (obj[h] !== undefined) cur[j] = safe_(obj[h]); });
  sh.getRange(row, 1, 1, head.length).setValues([cur]);
}

function settings_() {
  var s = {};
  rows_('Settings').forEach(function (r) { s[String(r.key).trim()] = r.value; });
  return s;
}

function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function dom_(v) { return str_(v).toLowerCase().replace(/@.*$/, ''); }

function parseDate_(v) {
  if (v instanceof Date) return v;
  var s = str_(v);
  if (!s) return null;
  if (/T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})$/.test(s)) { var dz = new Date(s); return isNaN(dz) ? null : dz; }
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], (+(m[4] || 0)) - 7, +(m[5] || 0)));
}

function fmtDate_(d) {
  if (!d) return '';
  var v = new Date(d.getTime() + 7 * 3600 * 1000);
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  return p(v.getUTCDate()) + '/' + p(v.getUTCMonth() + 1) + ' ' + p(v.getUTCHours()) + ':' + p(v.getUTCMinutes());
}

function now_() {
  if (typeof DEMO_NOW !== 'undefined' && DEMO_NOW) return new Date(DEMO_NOW);
  return new Date();
}

function phase_(s) {
  if (TEST_MODE) return 'test';
  var n = now_().getTime();
  var ns = parseDate_(s.nominationStart), ne = parseDate_(s.nominationEnd);
  var vs = parseDate_(s.votingStart), ve = parseDate_(s.votingEnd);
  if (ns && n < ns.getTime()) return 'upcoming';
  if (ne && n <= ne.getTime() + 59999) return 'nomination';
  if (vs && n < vs.getTime()) return 'review';
  if (ve && n <= ve.getTime() + 59999) return 'voting';
  return 'closed';
}

function cycle_(s) {
  return {
    cycleName: str_(s.cycleName), prize: str_(s.prize),
    nominationStart: fmtDate_(parseDate_(s.nominationStart)), nominationEnd: fmtDate_(parseDate_(s.nominationEnd)),
    votingStart: fmtDate_(parseDate_(s.votingStart)), votingEnd: fmtDate_(parseDate_(s.votingEnd)),
    finalizeDate: str_(s.finalizeDate), ceremonyDate: str_(s.ceremonyDate),
    phase: phase_(s), testMode: TEST_MODE
  };
}

function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function iso_() { return now_().toISOString(); }

/* ---------------- Employees & auth ---------------- */

/** role trong sheet: employee | leader | sub-leader | admin (không phân biệt hoa thường). */
function roleOf_(v) {
  var r = str_(v).toLowerCase();
  if (r.indexOf('admin') >= 0 || r === 'hr') return 'admin';
  if (r.indexOf('sub') >= 0) return 'subleader';
  if (r.indexOf('lead') >= 0 || r === 'tl') return 'leader';
  return 'employee';
}

function canLogin_(role) { return role === 'leader' || role === 'subleader' || role === 'admin'; }

var EMP_CACHE_ = null;
/** Cột unit: nhiều team cách nhau bằng dấu ;  Tên team có dấu * ở cuối = chỉ là nhân viên của team đó (không phụ trách).
 *  VD "Sales VuLB; Sales Management & Agency*": phụ trách Sales VuLB, là nhân viên của Sales Management & Agency. */
function parseUnits_(v) {
  var out = [], seen = {};
  String(v == null ? '' : v).split(/[;；]/).forEach(function (x) {
    x = x.trim(); var member = /\*$/.test(x); x = x.replace(/\*+$/, '').trim();
    if (!x) return;
    if (seen[x] !== undefined) { if (!member) out[seen[x]].member = false; return; }
    seen[x] = out.length; out.push({ n: x, member: member });
  });
  return out;
}
function units_(v) { return parseUnits_(v).map(function (x) { return x.n; }); }
function leadUnits_(v) { return parseUnits_(v).filter(function (x) { return !x.member; }).map(function (x) { return x.n; }); }
function unitStr_(v) { return parseUnits_(v).map(function (x) { return x.n + (x.member ? '*' : ''); }).join('; '); }
function overlap_(a, b) { return (a || []).some(function (x) { return (b || []).indexOf(x) >= 0; }); }

function employees_() {
  if (EMP_CACHE_) return EMP_CACHE_;
  var map = {};
  rows_('Employees').forEach(function (e) {
    var d = dom_(e.domain);
    if (d) map[d] = { domain: d, name: str_(e.name), title: str_(e.title), unit: unitStr_(e.unit), units: units_(e.unit), leads: leadUnits_(e.unit), role: roleOf_(e.role),
      accessCode: str_(e.accessCode), active: String(e.active).toUpperCase() !== 'FALSE', _row: e._row };
  });
  EMP_CACHE_ = map;
  return map;
}

function auth_(req) {
  var d = dom_(req.domain);
  var code = str_(req.code).toUpperCase();
  if (!d || !code) throw new Error('Vui lòng nhập domain và mã đăng nhập.');
  var u = employees_()[d];
  if (!u || !canLogin_(u.role) || !u.active || u.accessCode.toUpperCase() !== code) {
    throw new Error('Domain hoặc mã đăng nhập không đúng.');
  }
  return { domain: u.domain, name: u.name, role: u.role === 'admin' ? 'admin' : 'leader', leaderType: u.role, unit: u.leads.join('; ') || u.units.join('; '), units: u.units, leads: u.leads, block: voteBlock_(u.domain), title: u.title };
}

function requireAdmin_(me) { if (me.role !== 'admin') throw new Error('Chỉ Admin được dùng chức năng này.'); }
function requireLeader_(me) { if (me.role !== 'leader') throw new Error('Chức năng dành cho Team Leader / Sub-team Leader.'); }

function generateCodes_() {
  var alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  var used = {};
  var list = rows_('Employees');
  list.forEach(function (u) { if (str_(u.accessCode)) used[str_(u.accessCode).toUpperCase()] = 1; });
  var n = 0;
  list.forEach(function (u) {
    if (dom_(u.domain) && canLogin_(roleOf_(u.role)) && !str_(u.accessCode)) {
      var c;
      do { c = ''; for (var i = 0; i < 8; i++) c += alphabet.charAt(Math.floor(Math.random() * alphabet.length)); } while (used[c]);
      used[c] = 1;
      update_('Employees', u._row, { accessCode: c });
      n++;
    }
  });
  return n;
}

/* ---------------- Domain helpers ---------------- */

function quotaFor_(headcount) {
  return Math.max(1, Math.floor((Number(headcount) || 0) * 0.1));
}

/** Team/sub-team và headcount (FTE + CTV) đếm tự động từ tab Employees; HR (admin) không tính. */
function teams_(emp) {
  emp = emp || employees_();
  var map = {};
  Object.keys(emp).forEach(function (k) {
    var e = emp[k];
    if (!e.active || e.role === 'admin') return;
    e.units.forEach(function (un) {
      if (!map[un]) map[un] = { unit: un, headcount: 0 };
      map[un].headcount++;
    });
  });
  return map;
}

/** Đề cử hiển thị ngay khi gửi (approved). HR có thể ẩn (hidden) – đề cử bị ẩn không tính vào quota. */
function visible_(n) { return n.status === 'approved' || n.status === 'pending'; }
function activeNomination_(n) { return visible_(n); }

/** Thông tin ứng viên: tên/title lấy theo domain từ tab Employees (cập nhật mới nhất), fallback bản lưu lúc đề cử. */
function candidate_(n, emp) {
  var e = emp[dom_(n.candidateDomain)] || {};
  return {
    id: str_(n.id), candidateDomain: dom_(n.candidateDomain), candidateName: e.name || str_(n.candidateName),
    title: e.title || str_(n.candidateTitle), unit: str_(n.unit), cunits: e.units || [], sow: str_(n.sow),
    ev1: str_(n.ev1), ev2: str_(n.ev2), ev3: str_(n.ev3), ev4: str_(n.ev4),
    r1: Number(n.r1) || 0, r2: Number(n.r2) || 0, r3: Number(n.r3) || 0, r4: Number(n.r4) || 0
  };
}

var INCLUDE_SELF_ = false;

function scoreOf_(v) {
  return CRITERIA.reduce(function (s, c) { return s + (Number(v[c.key]) / 5) * c.weight; }, 0);
}

function round1_(x) { return Math.round(x * 10) / 10; }

/**
 * Xếp hạng: 1) số lượt vote nhiều hơn; 2) bằng nhau → điểm trung bình cao hơn; 3) vẫn bằng → điểm Impact cao hơn.
 * includeSelf = true: tính thêm điểm tự đánh giá của leader đề cử (r1..r4) như 1 phiếu khi lấy trung bình (không tính vào số lượt vote).
 */
function ranking_(s, noms, votes, emp, includeSelf) {
  var minJ = Number(s.minJudges) || 2;
  var list = noms.filter(visible_).map(function (n) {
    var c = candidate_(n, emp);
    var vs = votes.filter(function (v) { return str_(v.nominationId) === str_(n.id); }).map(function (v) {
      return { c1: Number(v.c1), c2: Number(v.c2), c3: Number(v.c3), c4: Number(v.c4) };
    });
    var k = vs.length;
    var self = { c1: Number(n.r1) || 0, c2: Number(n.r2) || 0, c3: Number(n.r3) || 0, c4: Number(n.r4) || 0 };
    var hasSelf = self.c1 && self.c2 && self.c3 && self.c4;
    var pool = vs.slice();
    if (includeSelf && hasSelf) pool.push(self);
    var m = pool.length;
    var avg = {};
    CRITERIA.forEach(function (cr) { avg[cr.key] = m ? round1_(pool.reduce(function (a, v) { return a + v[cr.key]; }, 0) / m) : 0; });
    var mean = m ? pool.map(scoreOf_).reduce(function (a, b) { return a + b; }, 0) / m : 0;
    return {
      id: c.id, candidateName: c.candidateName, candidateDomain: c.candidateDomain, title: c.title, unit: c.unit,
      judges: k, avg: avg, impactRaw: m ? pool.reduce(function (a, v) { return a + v.c2; }, 0) / m : 0,
      score: round1_(mean), selfScore: hasSelf ? round1_(scoreOf_(self)) : null, eligible: k >= minJ
    };
  });
  list.sort(function (a, b) {
    if (b.judges !== a.judges) return b.judges - a.judges;
    if (b.score !== a.score) return b.score - a.score;
    return b.impactRaw - a.impactRaw;
  });
  list.forEach(function (r, i) { r.rank = i + 1; delete r.impactRaw; });
  return { method: 'votes-then-average', includeSelf: !!includeSelf, minJudges: minJ, list: list };
}

/* ---------------- Public ---------------- */

function publicData_() {
  var s = settings_();
  var emp = employees_();
  return {
    cycle: cycle_(s),
    candidates: rows_('Nominations').filter(visible_).map(function (n) { return candidate_(n, emp); })
  };
}

/* ---------------- Leader ---------------- */

function home_(me) {
  var s = settings_();
  var res = { me: me, cycle: cycle_(s) };
  if (me.role === 'leader') {
    var emp = employees_();
    var tm = teams_(emp), all = rows_('Nominations');
    var myUnits = me.leads && me.leads.length ? me.leads : [me.unit];
    res.teams = myUnits.map(function (un) {
      var t = tm[un] || { unit: un, headcount: 0 };
      return { unit: un, headcount: t.headcount, quotaMax: quotaFor_(t.headcount),
        used: all.filter(function (n) { return str_(n.unit) === un && activeNomination_(n); }).length };
    });
    res.team = res.teams[0];
    var mine = all.filter(function (n) { return myUnits.indexOf(str_(n.unit)) >= 0; });
    res.members = Object.keys(emp).map(function (k) { return emp[k]; })
      .filter(function (e) { return e.active && overlap_(e.units, myUnits); })
      .map(function (e) { return { domain: e.domain, name: e.name, title: e.title, units: e.units }; });
    res.nominations = mine.map(function (n) {
      var o = candidate_(n, emp);
      o.status = str_(n.status); o.hrNote = str_(n.hrNote); o.createdAt = str_(n.createdAt); o.nominatorName = str_(n.nominatorName);
      return o;
    });
    res.hasVoted = liveVotes_().some(function (v) { return dom_(v.voterDomain) === me.domain; });
  }
  return res;
}

function clean_(v, max, label, required) {
  var s = str_(v);
  if (required && !s) throw new Error('Vui lòng nhập ' + label + '.');
  if (s.length > max) throw new Error(label + ' tối đa ' + max + ' ký tự.');
  return s;
}

function submitNomination_(me, d) {
  requireLeader_(me);
  var s = settings_();
  if (['nomination', 'test'].indexOf(phase_(s)) < 0) throw new Error('Hiện không trong thời gian đề cử.');
  var myUnits = me.leads && me.leads.length ? me.leads : [me.unit];
  var un = str_(d.unit) || myUnits[0];
  if (myUnits.indexOf(un) < 0) throw new Error('Bạn không phụ trách team "' + un + '".');
  var t = teams_()[un];
  if (!t) throw new Error('Team "' + un + '" chưa có trong danh sách nhân viên. Liên hệ HR.');
  var noms = rows_('Nominations');
  var used = noms.filter(function (n) { return str_(n.unit) === un && activeNomination_(n); }).length;
  var max = quotaFor_(t.headcount);
  if (used >= max) throw new Error('Team đã dùng hết quota đề cử (' + max + '). Rút một đề cử đang chờ duyệt để gửi lại.');
  var cd = dom_(d.candidateDomain);
  if (!cd) throw new Error('Vui lòng nhập domain của người bạn muốn đề cử.');
  var c = employees_()[cd];
  if (!c || !c.active) throw new Error('Không tìm thấy domain "' + cd + '" trong danh sách nhân viên. Liên hệ HR nếu danh sách thiếu.');
  if (c.units.indexOf(un) < 0) throw new Error(c.name + ' thuộc team ' + c.unit + '. Bạn chỉ đề cử được thành viên team ' + un + '.');
  if (noms.some(function (n) { return dom_(n.candidateDomain) === cd && activeNomination_(n); })) throw new Error(c.name + ' đã được đề cử trong đợt này.');
  var rec = {
    id: uid_(), createdAt: iso_(), nominatorDomain: me.domain, nominatorName: me.name, unit: un,
    candidateDomain: cd, candidateName: c.name, candidateTitle: c.title,
    sow: clean_(d.sow, LIMITS.sow, 'SOW', true), status: 'approved', hrNote: '', updatedAt: iso_()
  };
  CRITERIA.forEach(function (cr) {
    rec[cr.ev] = clean_(d[cr.ev], LIMITS.ev, 'minh chứng cho tiêu chí ' + cr.name, true);
    var x = Number(d[cr.r]);
    if (!(x >= 1 && x <= 5 && Math.floor(x) === x)) throw new Error('Vui lòng chọn mức đánh giá (1–5) cho tiêu chí ' + cr.name + '.');
    rec[cr.r] = x;
  });
  append_('Nominations', rec);
  return home_(me);
}

function withdrawNomination_(me, id) {
  requireLeader_(me);
  if (['nomination', 'test'].indexOf(phase_(settings_())) < 0) throw new Error('Đã hết thời gian đề cử.');
  var n = rows_('Nominations').filter(function (x) { return str_(x.id) === str_(id); })[0];
  if (!n || (me.leads && me.leads.length ? me.leads : [me.unit]).indexOf(str_(n.unit)) < 0) throw new Error('Không tìm thấy đề cử.');
  if (n.status === 'withdrawn') throw new Error('Đề cử đã được rút.');
  update_('Nominations', n._row, { status: 'withdrawn', updatedAt: iso_() });
  return home_(me);
}

/** Ứng viên thuộc team của người chấm: team đề cử thuộc team của họ, hoặc bản thân ứng viên có team trùng. */
function ownsCand_(me, c, emp) {
  var mu = me.units && me.units.length ? me.units : [me.unit];
  var e = emp[c.candidateDomain];
  if (mu.indexOf(c.unit) >= 0 || overlap_(mu, e && e.units)) return true;
  var bl = me.block || [];
  return [c.unit].concat((e && e.units) || []).some(function (un) {
    un = String(un).toLowerCase();
    return bl.some(function (pre) { return un.indexOf(pre) === 0; });
  });
}

function ballot_(me) {
  requireLeader_(me);
  var emp = employees_();
  var mine = liveVotes_().filter(function (v) { return dom_(v.voterDomain) === me.domain; });
  return {
    cycle: cycle_(settings_()),
    candidates: rows_('Nominations').filter(visible_).map(function (n) {
      var o = candidate_(n, emp); o.locked = ownsCand_(me, o, emp); return o;
    }),
    myVotes: mine.map(function (v) {
      return { nominationId: str_(v.nominationId), candidateName: str_(v.candidateName), candidateUnit: str_(v.candidateUnit),
        c1: Number(v.c1), c2: Number(v.c2), c3: Number(v.c3), c4: Number(v.c4), total: Number(v.total), comment: str_(v.comment), createdAt: str_(v.createdAt) };
    })
  };
}

function submitVote_(me, entries) {
  requireLeader_(me);
  if (['voting', 'test'].indexOf(phase_(settings_())) < 0) throw new Error('Hiện không trong thời gian chấm điểm.');
  var mine = liveVotes_().filter(function (v) { return dom_(v.voterDomain) === me.domain; });
  if (!Array.isArray(entries) || entries.length < 1) throw new Error('Chọn ứng viên để chấm.');
  if (mine.length + entries.length > 2) throw new Error(mine.length >= 2 ? 'Bạn đã dùng hết 2 lượt chấm.' : 'Bạn chỉ còn ' + (2 - mine.length) + ' lượt chấm.');
  var noms = rows_('Nominations');
  var emp = employees_();
  var seen = {};
  mine.forEach(function (v) { seen[str_(v.nominationId)] = 1; });
  var recs = entries.map(function (e) {
    var id = str_(e.nominationId);
    if (seen[id]) throw new Error('Bạn đã chấm ứng viên này rồi. Mỗi ứng viên chỉ chấm 1 lần.');
    seen[id] = 1;
    var n = noms.filter(function (x) { return str_(x.id) === id && visible_(x); })[0];
    if (!n) throw new Error('Ứng viên không hợp lệ.');
    var c = candidate_(n, emp);
    if (ownsCand_(me, c, emp)) throw new Error('Không được chấm ứng viên thuộc team của bạn.');
    var sc = {};
    CRITERIA.forEach(function (cr) {
      var x = Number(e[cr.key]);
      if (!(x >= 1 && x <= 5 && Math.floor(x) === x)) throw new Error('Chấm đủ 4 tiêu chí (1–5) cho ' + c.candidateName + '.');
      sc[cr.key] = x;
    });
    return {
      id: uid_(), createdAt: iso_(), voterDomain: me.domain, voterName: me.name, voterUnit: me.unit,
      nominationId: id, candidateDomain: c.candidateDomain, candidateName: c.candidateName, candidateUnit: c.unit,
      c1: sc.c1, c2: sc.c2, c3: sc.c3, c4: sc.c4, total: round1_(scoreOf_(sc)),
      comment: clean_(e.comment, LIMITS.comment, 'lý do chấm điểm cho ' + c.candidateName, true)
    };
  });
  recs.forEach(function (r) { append_('Votes', r); });
  return ballot_(me);
}

/* ---------------- Admin ---------------- */

function adminData_() {
  var s = settings_();
  var emp = employees_();
  var noms = rows_('Nominations');
  var votes = liveVotes_(noms);
  var tmap = teams_(emp);
  var votedBy = {};
  votes.forEach(function (v) { votedBy[dom_(v.voterDomain)] = 1; });
  var people = Object.keys(emp).map(function (k) { return emp[k]; });
  var leaders = people.filter(function (u) { return u.role === 'leader' || u.role === 'subleader'; });
  Object.keys(tmap).forEach(function (k) { tmap[k].leaders = 0; tmap[k].voted = 0; });
  leaders.forEach(function (u) {
    u.leads.forEach(function (un) {
      if (!tmap[un]) tmap[un] = { unit: un, headcount: 0, leaders: 0, voted: 0 };
      tmap[un].leaders++;
      if (votedBy[u.domain]) tmap[un].voted++;
    });
  });
  var teams = Object.keys(tmap).map(function (k) {
    var t = tmap[k];
    var mine = noms.filter(function (n) { return str_(n.unit) === k; });
    t.quotaMax = quotaFor_(t.headcount);
    t.nominated = mine.filter(activeNomination_).length;
    t.hidden = mine.filter(function (n) { return n.status === 'hidden'; }).length;
    t.overQuota = t.nominated > t.quotaMax;
    return t;
  });
  var rk = ranking_(s, noms, votes, emp, INCLUDE_SELF_);
  return {
    cycle: cycle_(s),
    settingsRaw: { rankingMethod: rk.method, includeSelf: rk.includeSelf, minJudges: rk.minJudges },
    teams: teams.sort(function (a, b) { return a.unit.localeCompare(b.unit); }),
    employeeCount: people.length,
    employees: people.reduce(function (acc, u) {
      (u.units.length ? u.units : ['']).forEach(function (un) { acc.push({ domain: u.domain, name: u.name, title: u.title, unit: un, unitAll: u.unit, role: u.role, active: u.active }); });
      return acc;
    }, []),
    users: people.filter(function (u) { return canLogin_(u.role); }).map(function (u) {
      return { domain: u.domain, name: u.name, role: u.role, unit: u.unit, title: u.title, accessCode: u.accessCode, active: u.active, voted: !!votedBy[u.domain] };
    }),
    nominations: noms.filter(function (n) { return n.status !== 'deleted'; }).map(function (n) {
      var o = candidate_(n, emp);
      o.status = str_(n.status); o.hrNote = str_(n.hrNote); o.createdAt = str_(n.createdAt);
      o.nominatorName = str_(n.nominatorName); o.nominatorDomain = dom_(n.nominatorDomain);
      return o;
    }),
    votes: votes.filter(function (v) { return noms.some(function (n) { return str_(n.id) === str_(v.nominationId) && n.status !== 'deleted'; }); }).map(function (v) {
      return { voterDomain: dom_(v.voterDomain), voterName: str_(v.voterName), voterUnit: str_(v.voterUnit), nominationId: str_(v.nominationId),
        candidateName: str_(v.candidateName), candidateUnit: str_(v.candidateUnit), c1: Number(v.c1), c2: Number(v.c2), c3: Number(v.c3), c4: Number(v.c4),
        total: Number(v.total), comment: str_(v.comment), createdAt: str_(v.createdAt) };
    }),
    ranking: rk.list
  };
}

function adminReview_(id, status, note) {
  if (['approved', 'hidden'].indexOf(status) < 0) throw new Error('Trạng thái không hợp lệ.');
  var n = rows_('Nominations').filter(function (x) { return str_(x.id) === str_(id); })[0];
  if (!n) throw new Error('Không tìm thấy đề cử.');
  if (n.status === 'withdrawn') throw new Error('Đề cử đã được leader rút.');
  var upd = { status: status, updatedAt: iso_() };
  if (note !== undefined) upd.hrNote = clean_(note, LIMITS.note, 'Ghi chú', false);
  update_('Nominations', n._row, upd);
  return adminData_();
}

function adminEdit_(id, d) {
  var n = rows_('Nominations').filter(function (x) { return str_(x.id) === str_(id); })[0];
  if (!n) throw new Error('Không tìm thấy đề cử.');
  var upd = { updatedAt: iso_() };
  if (d.sow !== undefined) upd.sow = clean_(d.sow, LIMITS.sow, 'SOW', true);
  CRITERIA.forEach(function (cr) {
    if (d[cr.ev] !== undefined) upd[cr.ev] = clean_(d[cr.ev], LIMITS.ev, 'minh chứng ' + cr.name, true);
    if (d[cr.r] !== undefined) { var x = Number(d[cr.r]); if (!(x >= 1 && x <= 5 && Math.floor(x) === x)) throw new Error('Mức tự đánh giá phải từ 1–5.'); upd[cr.r] = x; }
  });
  update_('Nominations', n._row, upd);
  return adminData_();
}

/* ---------------- Import danh sách nhân viên ---------------- */

/**
 * rows: [{domain, name, title, unit, role}]. mode 'replace' = thay toàn bộ danh sách, 'merge' = thêm mới / cập nhật theo domain.
 * Mã đăng nhập cũ của cùng domain được giữ nguyên; mã mới tự tạo cho leader / sub-leader / admin.
 */
function adminImport_(rows, mode) {
  if (!Array.isArray(rows) || !rows.length) throw new Error('Danh sách trống.');
  if (rows.length > 5000) throw new Error('Tối đa 5.000 dòng mỗi lần import.');
  var old = employees_();
  var out = {}, order = [], errors = [];
  if (mode === 'merge') Object.keys(old).forEach(function (k) { out[k] = old[k]; order.push(k); });
  rows.forEach(function (r, i) {
    var d = dom_(r.domain), name = str_(r.name), unit = unitStr_(r.unit);
    if (!d && !name) return;
    if (!/^[a-z0-9._-]{2,64}$/.test(d)) { errors.push('Dòng ' + (i + 1) + ': domain "' + str_(r.domain) + '" không hợp lệ'); return; }
    if (!name || !unit) { errors.push('Dòng ' + (i + 1) + ' (' + d + '): thiếu họ tên hoặc team'); return; }
    var prev = old[d] || {};
    if (!out[d]) order.push(d);
    out[d] = { domain: d, name: name.slice(0, 120), title: str_(r.title).slice(0, 120), unit: unit.slice(0, 120), role: roleOf_(r.role),
      accessCode: prev.accessCode || '', active: true };
  });
  if (errors.length) throw new Error('Chưa import. Sửa các lỗi sau rồi thử lại: ' + errors.slice(0, 8).join('; ') + (errors.length > 8 ? '; …' : ''));
  if (!order.length) throw new Error('Không có dòng hợp lệ để import.');
  var roleName = { admin: 'admin', leader: 'leader', subleader: 'sub-leader', employee: 'employee' };
  var values = [SHEETS.Employees].concat(order.map(function (k) {
    var e = out[k];
    return [e.domain, safe_(e.name), safe_(e.title), safe_(e.unit), roleName[e.role] || 'employee', e.accessCode, e.active ? 'TRUE' : 'FALSE'];
  }));
  var sh = sheet_('Employees');
  EMP_CACHE_ = null;
  cacheDrop_('Employees');
  sh.clearContents();
  sh.getRange(1, 1, values.length, SHEETS.Employees.length).setValues(values);
  var created = generateCodes_();
  EMP_CACHE_ = null;
  var res = adminData_();
  res.importSummary = { total: order.length, created: created };
  return res;
}

/** Sửa 1 nhân viên (khi import đọc sai). Giữ nguyên mã đăng nhập; tự tạo mã nếu đổi sang leader/admin. */
function adminUpdateEmployee_(me, domainKey, d) {
  var emp = employees_();
  var cur = emp[dom_(domainKey)];
  if (!cur) throw new Error('Không tìm thấy nhân viên.');
  var nd = dom_(d.domain !== undefined ? d.domain : cur.domain);
  if (!/^[a-z0-9._-]{2,64}$/.test(nd)) throw new Error('Domain không hợp lệ.');
  if (nd !== cur.domain && emp[nd]) throw new Error('Domain ' + nd + ' đã tồn tại.');
  var name = str_(d.name !== undefined ? d.name : cur.name), unit = unitStr_(d.unit !== undefined ? d.unit : cur.unit);
  if (!name || !unit) throw new Error('Họ tên và team không được trống.');
  var role = d.role !== undefined ? roleOf_(d.role) : cur.role;
  var active = d.active !== undefined ? !!d.active : cur.active;
  if (cur.domain === me.domain && (role !== 'admin' || !active)) throw new Error('Không thể bỏ quyền Admin hoặc khoá chính tài khoản đang đăng nhập.');
  var roleName = { admin: 'admin', leader: 'leader', subleader: 'sub-leader', employee: 'employee' };
  update_('Employees', cur._row, { domain: nd, name: name.slice(0, 120), title: str_(d.title !== undefined ? d.title : cur.title).slice(0, 120),
    unit: unit.slice(0, 120), role: roleName[role], active: active ? 'TRUE' : 'FALSE' });
  generateCodes_();
  return adminData_();
}

/** Xoá đề cử: ẩn khỏi mọi nơi (status = deleted), trả lại quota cho team; phiếu chấm của đề cử này không còn được tính. */
function adminDelete_(id) {
  var n = rows_('Nominations').filter(function (x) { return str_(x.id) === str_(id); })[0];
  if (!n || n.status === 'deleted') throw new Error('Không tìm thấy đề cử.');
  update_('Nominations', n._row, { status: 'deleted', updatedAt: iso_() });
  return adminData_();
}

/** Toàn bộ dữ liệu leader cần trong 1 lần gọi (đề cử của team + danh sách đề cử + phiếu chấm). */
function leaderAll_(me) {
  var res = home_(me);
  res.pub = publicData_();
  res.ballot = ballot_(me);
  return res;
}
/* Dữ liệu minh hoạ cho chế độ DEMO + API giả lập. Tên người, domain là giả định. */
function __seed(withSample) {
  __DB = {};
  setup();
  var E = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  [
    ['hr.admin', 'Admin', 'HRBP', 'People & Culture', 'admin', 'Q7MX4KTB', 'TRUE'],
    ['binhnv', 'Nguyễn Văn Bình', 'Team Leader', 'Planner', 'leader', 'R2HV9PDN', 'TRUE'],
    ['cucntt', 'Trần Thị Cúc', 'Sub-team Leader', 'Sales VuLB', 'sub-leader', 'W8KZ3MFA', 'TRUE'],
    ['dunglh', 'Lê Hoàng Dũng', 'Team Leader', 'Marketing', 'leader', 'T5NB7QYC', 'TRUE'],
    ['vinhnq5', 'Nguyễn Quang Vinh', 'Team Leader', 'Ad Commercial; Sales VuLB', 'leader', 'V4NQ6HXD', 'TRUE'],
    ['empm', 'Phạm Mai Em', 'Team Leader', 'Audience Insights', 'leader', 'J3XR8WDK', 'TRUE'],
    ['phongdq', 'Đặng Quốc Phong', 'Team Leader', 'Ad Commercial', 'leader', 'M6PC2VHT', 'TRUE'],
    ['annt', 'Nguyễn Thị An', 'Senior Account Manager', 'Sales VuLB', 'employee', '', 'TRUE'],
    ['khoalm', 'Lê Minh Khoa', 'Ad Operations Specialist', 'Ad Commercial', 'employee', '', 'TRUE'],
    ['hapt', 'Phạm Thu Hà', 'Media Planner', 'Planner', 'employee', '', 'TRUE'],
    ['tuanlq', 'Lý Quốc Tuấn', 'Senior Planner', 'Planner', 'employee', '', 'TRUE'],
    ['huydq', 'Đỗ Quang Huy', 'Data Analyst', 'Audience Insights', 'employee', '', 'TRUE'],
    ['lanvn', 'Vũ Ngọc Lan', 'Content Lead', 'Content Commercial', 'employee', '', 'TRUE'],
    ['baohg', 'Hoàng Gia Bảo', 'Marketing Executive', 'Marketing', 'employee', '', 'TRUE']
  ].forEach(function (r) { E.appendRow(r); });
  var hc = { 'Planner': 24, 'Sales VuLB': 15, 'Marketing': 8, 'Audience Insights': 12, 'Ad Commercial': 21, 'Content Commercial': 18 };
  var cur = {}; rows_('Employees').forEach(function (e) { if (e.role !== 'admin') cur[e.unit] = (cur[e.unit] || 0) + 1; });
  Object.keys(hc).forEach(function (u, ui) { for (var i = (cur[u] || 0); i < hc[u]; i++) E.appendRow(['nv' + ui + '_' + i, 'Nhân viên ' + u + ' ' + (i + 1), 'Executive', u, 'employee', '', 'TRUE']); });
  if (!withSample) { __save(); return; }
  var noms = [
    ['annt', 'Nguyễn Thị An', 'Senior Account Manager', 'Sales VuLB', 'Phụ trách nhóm khách hàng FMCG; chịu trách nhiệm doanh thu và tái ký hợp đồng trong quý.', 'Đạt 118% KPI doanh thu Q4; giữ tỷ lệ đúng hạn báo giá 100%.', 'Tái ký 3 khách hàng lớn, đóng góp phần lớn doanh thu mảng FMCG của sub-team.', 'Chủ động hỗ trợ 2 bạn mới hội nhập, chia sẻ playbook đàm phán.', 'Dùng AI tóm tắt brief khách hàng, rút ngắn thời gian chuẩn bị proposal.', 'approved'],
    ['khoalm', 'Lê Minh Khoa', 'Ad Operations Specialist', 'Ad Commercial', 'Vận hành chiến dịch: setup, monitoring, báo cáo; đảm bảo SLA booking.', 'Không trễ SLA booking nào trong quý.', 'Giảm lỗi setup chiến dịch nhờ chuẩn hoá quy trình QA.', 'Hỗ trợ team Sales xử lý sự cố ngoài giờ.', 'Xây checklist QA tự động trên Google Sheet.', 'approved'],
    ['hapt', 'Phạm Thu Hà', 'Media Planner', 'Planner', 'Lập kế hoạch truyền thông cho nhóm khách hàng ngành Sữa.', '5 proposal thắng pitching trong quý.', 'Đóng góp vào 2 hợp đồng mới ngành Sữa.', 'Chủ động đào tạo nội bộ về công cụ planning.', 'Xây template plan dùng chung cho team.', 'approved'],
    ['huydq', 'Đỗ Quang Huy', 'Data Analyst', 'Audience Insights', 'Phân tích dữ liệu audience, hỗ trợ đề xuất targeting.', 'Hoàn thành 100% yêu cầu insight đúng hạn.', 'Insight giúp Sales chốt 3 deal mới.', 'Tự đề xuất dashboard theo dõi chung.', 'Ứng dụng AI để tự động hoá báo cáo hằng tuần.', 'approved'],
    ['lanvn', 'Vũ Ngọc Lan', 'Content Lead', 'Content Commercial', 'Phụ trách nội dung quảng cáo thương mại cho khách hàng chiến lược.', '12 bộ nội dung đạt KPI tương tác.', 'Nội dung được khách hàng gia hạn hợp đồng.', 'Kèm cặp 3 bạn content mới.', 'Áp dụng AI để rút ngắn thời gian sản xuất.', 'approved'],
    ['baohg', 'Hoàng Gia Bảo', 'Marketing Executive', 'Marketing', 'Triển khai hoạt động marketing thương hiệu Adtima.', 'Tổ chức 2 sự kiện khách hàng đúng kế hoạch.', 'Tăng lượt đăng ký webinar.', 'Chủ động phối hợp với Sales.', 'Thử nghiệm kênh mới.', 'approved']
  ];
  noms.forEach(function (n, i) {
    append_('Nominations', { id: 'demo' + i, createdAt: '2026-11-10T0' + (i + 2) + ':00:00.000Z', nominatorDomain: 'demo', nominatorName: 'Leader ' + n[3], unit: n[3], candidateDomain: n[0], candidateName: n[1], candidateTitle: n[2], sow: n[4], ev1: n[5], ev2: n[6], ev3: n[7], ev4: n[8], r1: [5,4,4,5,4,4][i], r2: [4,4,5,4,5,4][i], r3: [4,5,4,4,4,4][i], r4: [4,4,4,5,4,3][i], status: n[9], hrNote: '', updatedAt: '' });
  });
  var votes = [
    ['dunglh', 'Lê Hoàng Dũng', 'Marketing', 'demo0', 5, 4, 4, 3], ['dunglh', 'Lê Hoàng Dũng', 'Marketing', 'demo1', 4, 5, 3, 4],
    ['empm', 'Phạm Mai Em', 'Audience Insights', 'demo0', 4, 4, 4, 4], ['empm', 'Phạm Mai Em', 'Audience Insights', 'demo2', 5, 5, 4, 4],
    ['phongdq', 'Đặng Quốc Phong', 'Ad Commercial', 'demo2', 4, 4, 5, 4], ['phongdq', 'Đặng Quốc Phong', 'Ad Commercial', 'demo3', 5, 4, 4, 5]
  ];
  var nomRows = rows_('Nominations');
  votes.forEach(function (v) {
    var n = nomRows.filter(function (x) { return x.id === v[3]; })[0];
    var sc = { c1: v[4], c2: v[5], c3: v[6], c4: v[7] };
    append_('Votes', { id: uid_(), createdAt: '2026-11-13T03:00:00.000Z', voterDomain: v[0], voterName: v[1], voterUnit: v[2], nominationId: v[3], candidateDomain: n.candidateDomain, candidateName: n.candidateName, candidateUnit: n.unit, c1: sc.c1, c2: sc.c2, c3: sc.c3, c4: sc.c4, total: round1_(scoreOf_(sc)), comment: 'Minh chứng rõ ràng, tác động thực tế tới khách hàng.' });
  });
  __save();
}
if (!__DB.Employees) { __seed(true); try { if (!__LS.getItem('eoq_demo_now')) { __LS.setItem('eoq_demo_now', '2026-11-11T03:00:00Z'); DEMO_NOW = '2026-11-11T03:00:00Z'; } } catch (e) {} }
var MockAPI = {
  call: function (req) {
    var out = doPost({ postData: { contents: JSON.stringify(req) } });
    __save();
    return JSON.parse(out.content);
  },
  reset: function (withSample) { __seed(!!withSample); },
  setNow: function (iso) { DEMO_NOW = iso || ''; try { __LS.setItem('eoq_demo_now', DEMO_NOW); } catch (e) {} }
};
if (typeof window !== 'undefined') window.MockAPI = MockAPI;
