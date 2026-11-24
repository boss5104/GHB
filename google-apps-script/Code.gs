/**
 * GH Debt Planner — ตัวเชื่อม Google Sheet (Google Apps Script)
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 *  1) สร้าง Google Sheet ใหม่ แล้วไปที่เมนู "ส่วนขยาย" → "Apps Script"
 *  2) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดในไฟล์นี้ลงไป กดบันทึก
 *  3) กด "การทำให้ใช้งานได้" (Deploy) → "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 *     - ประเภท: เว็บแอป (Web app)
 *     - เรียกใช้ในฐานะ (Execute as): ฉัน (Me)
 *     - ผู้มีสิทธิ์เข้าถึง (Who has access): ทุกคน (Anyone)
 *  4) อนุญาตสิทธิ์ แล้วคัดลอก URL ที่ลงท้ายด้วย /exec ไปวางในหน้า "ตั้งค่า Google Sheet" ของแอป
 *
 * ชีตที่สร้างขึ้นอัตโนมัติ
 *  - ข้อมูลปัจจุบัน    : ยอดตั้งต้นและยอดผ่อนของรอบเดือนล่าสุด (อ่านง่าย)
 *  - ประวัติการชำระ    : ประวัติทุกงวดที่บันทึก (อ่านง่าย)
 *  - แผนประนอมหนี้     : รายละเอียดโครงการประนอมหนี้ (อ่านง่าย)
 *  - บันทึกการอัปเดต   : ทุกครั้งที่มีการแก้ไข จะบันทึกว่าเปลี่ยนอะไร จากอุปกรณ์ไหน
 *  - _state             : ข้อมูลระบบของแอป (ซ่อนไว้ — ห้ามแก้ไขด้วยมือ)
 */

var STATE_SHEET = '_state';
var CURRENT_SHEET = 'ข้อมูลปัจจุบัน';
var HISTORY_SHEET = 'ประวัติการชำระ';
var RESTRUCTURE_SHEET = 'แผนประนอมหนี้';
var LOG_SHEET = 'บันทึกการอัปเดต';
var CHUNK_SIZE = 40000;
var MAX_LOG_ROWS = 1000;

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'ping') {
    return json_({ ok: true, app: 'ghb-debt-planner', api: 1 });
  }
  var stored = readState_();
  return json_({
    ok: true,
    rev: stored.rev,
    updatedAt: stored.updatedAt,
    lastDevice: stored.device,
    state: stored.state,
    log: readLog_(20)
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'save' || !body.state || typeof body.state !== 'object') {
      return json_({ ok: false, error: 'bad_request' });
    }
    var prev = readState_();
    var rev = (prev.rev || 0) + 1;
    var updatedAt = new Date().toISOString();
    var device = String(body.device || 'ไม่ระบุอุปกรณ์').slice(0, 60);
    var changes = Array.isArray(body.changes) ? body.changes.map(String).slice(0, 200) : [];

    writeState_(body.state, rev, updatedAt, device);
    writeReadable_(body.state);
    appendLog_(updatedAt, device, changes);

    return json_({ ok: true, rev: rev, updatedAt: updatedAt, prevRev: prev.rev || 0 });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ป้องกันไม่ให้ Sheets ตีความข้อความเป็นสูตร/วันที่
function text_(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function readState_() {
  var sh = sheet_(STATE_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return { rev: 0, updatedAt: null, device: null, state: null };
  var head = sh.getRange(1, 1, 1, 3).getValues()[0];
  var raw = sh.getRange(2, 1, last - 1, 1).getValues()
    .map(function (r) { return String(r[0] || '').slice(1); })   // ตัดเครื่องหมาย ~ นำหน้า
    .join('');
  var state = null;
  try { state = raw ? JSON.parse(raw) : null; } catch (err) { state = null; }
  return {
    rev: Number(head[0]) || 0,
    updatedAt: head[1] ? String(head[1]).slice(1) : null,
    device: head[2] ? String(head[2]).slice(1) : null,
    state: state
  };
}

function writeState_(state, rev, updatedAt, device) {
  var sh = sheet_(STATE_SHEET);
  var str = JSON.stringify(state);
  var rows = [];
  for (var i = 0; i < str.length; i += CHUNK_SIZE) rows.push(['~' + str.slice(i, i + CHUNK_SIZE)]);
  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([[rev, '~' + updatedAt, '~' + device, 'ข้อมูลระบบของแอป — กรุณาอย่าแก้ไขชีตนี้']]);
  if (rows.length) sh.getRange(2, 1, rows.length, 1).setValues(rows);
  try { if (!sh.isSheetHidden()) sh.hideSheet(); } catch (err) { /* ชีตเดียวที่มองเห็นได้ ซ่อนไม่ได้ */ }
}

function num_(v) { var n = Number(v); return isFinite(n) ? n : 0; }

function writeTable_(name, rows, opts) {
  var sh = sheet_(name);
  sh.clearContents();
  sh.clearFormats();
  if (!rows.length) return sh;
  var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
  var data = rows.map(function (r) {
    var out = r.slice();
    while (out.length < width) out.push('');
    return out;
  });
  sh.getRange(1, 1, data.length, width).setValues(data);
  var header = sh.getRange(1, 1, 1, width);
  header.setFontWeight('bold').setBackground('#0e1a2e').setFontColor('#e2c48f');
  sh.setFrozenRows(1);
  if (data.length > 1 && opts && opts.numberCols) {
    opts.numberCols.forEach(function (c) {
      if (c <= width) sh.getRange(2, c, data.length - 1, 1).setNumberFormat('#,##0.00');
    });
  }
  sh.autoResizeColumns(1, width);
  return sh;
}

function writeReadable_(s) {
  var cur = s.current || {};
  var h = cur.house || {};
  var d = cur.dorm || {};

  writeTable_(CURRENT_SHEET, [
    ['หัวข้อ', 'บัญชีบ้าน (4901)', 'บัญชีหอ (6367)', 'รวม'],
    ['เดือนที่คำนวณ', text_(cur.date || ''), '', ''],
    ['เงินต้นคงเหลือตั้งต้น (บาท)', num_(h.principal), num_(d.principal), num_(h.principal) + num_(d.principal)],
    ['ดอกเบี้ยค้างตั้งต้น (บาท)', num_(h.interest), num_(d.interest), num_(h.interest) + num_(d.interest)],
    ['อัตราดอกเบี้ยต่อปี (%)', num_(h.rate), num_(d.rate), ''],
    ['ยอดจ่ายตัดต้น (บาท)', num_(h.payPrincipal), num_(d.payPrincipal), num_(h.payPrincipal) + num_(d.payPrincipal)],
    ['ยอดจ่ายตัดดอก (บาท)', num_(h.payInterest), num_(d.payInterest), num_(h.payInterest) + num_(d.payInterest)],
    ['เงินต้นคงเหลือหลังหักจ่าย (บาท)', Math.max(0, num_(h.principal) - num_(h.payPrincipal)), Math.max(0, num_(d.principal) - num_(d.payPrincipal)),
      Math.max(0, num_(h.principal) - num_(h.payPrincipal)) + Math.max(0, num_(d.principal) - num_(d.payPrincipal))]
  ], { numberCols: [2, 3, 4] });

  var hist = Array.isArray(s.history) ? s.history : [];
  var histRows = [[
    'รอบเดือน',
    'บ้าน: เงินต้นตั้งต้น', 'บ้าน: ดอกเบี้ยค้าง', 'บ้าน: อัตรา (%)', 'บ้าน: จ่ายต้น', 'บ้าน: จ่ายดอก',
    'หอ: เงินต้นตั้งต้น', 'หอ: ดอกเบี้ยค้าง', 'หอ: อัตรา (%)', 'หอ: จ่ายต้น', 'หอ: จ่ายดอก',
    'รวมจ่ายรอบนี้', 'เงินต้นคงเหลือรวม'
  ]];
  hist.forEach(function (x) {
    var a = x.house || {}, b = x.dorm || {};
    histRows.push([
      text_(x.date || ''),
      num_(a.principal), num_(a.interest), num_(a.rate), num_(a.payPrincipal), num_(a.payInterest),
      num_(b.principal), num_(b.interest), num_(b.rate), num_(b.payPrincipal), num_(b.payInterest),
      num_(a.payPrincipal) + num_(a.payInterest) + num_(b.payPrincipal) + num_(b.payInterest),
      (num_(a.principal) - num_(a.payPrincipal)) + (num_(b.principal) - num_(b.payPrincipal))
    ]);
  });
  writeTable_(HISTORY_SHEET, histRows, { numberCols: [2, 3, 5, 6, 7, 8, 10, 11, 12, 13] });

  var r = s.restructure || {};
  var rRows = [['บัญชี', 'ชื่อโครงการ', 'อัตราดอกเบี้ยพิเศษ (%)', 'ช่วงที่', 'เริ่ม (เดือน/ปี พ.ศ.)', 'ระยะเวลา (ปี)', 'ระยะเวลา (เดือน)', 'ค่างวด (บาท/เดือน)']];
  rRows.push(['สถานะโครงการ', r.enabled ? 'เปิดใช้งาน' : 'ปิดใช้งาน', '', '', '', '', '', '']);
  [['house', 'บ้าน (4901)'], ['dorm', 'หอ (6367)']].forEach(function (pair) {
    var acc = r[pair[0]] || {};
    var periods = Array.isArray(acc.periods) ? acc.periods : [];
    if (!periods.length) rRows.push([pair[1], text_(acc.name || ''), num_(acc.rate), '-', '', '', '', '']);
    periods.forEach(function (p, i) {
      rRows.push([pair[1], text_(acc.name || ''), num_(acc.rate), i + 1, text_(p.startMonth + '/' + p.startYear), num_(p.years), num_(p.months), num_(p.payment)]);
    });
  });
  writeTable_(RESTRUCTURE_SHEET, rRows, { numberCols: [8] });
}

function appendLog_(updatedAt, device, changes) {
  var sh = sheet_(LOG_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['เวลา', 'อุปกรณ์', 'รายการที่อัปเดต', 'จำนวนรายการ']])
      .setFontWeight('bold').setBackground('#0e1a2e').setFontColor('#e2c48f');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 170);
    sh.setColumnWidth(2, 140);
    sh.setColumnWidth(3, 520);
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 4).setValues([[new Date(updatedAt), text_(device), text_(changes.join('\n') || 'บันทึกข้อมูล'), changes.length]]);
  sh.getRange(2, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(2, 3).setWrap(true);
  var last = sh.getLastRow();
  if (last > MAX_LOG_ROWS + 1) sh.deleteRows(MAX_LOG_ROWS + 2, last - MAX_LOG_ROWS - 1);
}

function readLog_(n) {
  var sh = sheet_(LOG_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var count = Math.min(n, last - 1);
  return sh.getRange(2, 1, count, 3).getValues().map(function (r) {
    var t = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    return { t: t, device: String(r[1] || ''), changes: String(r[2] || '').split('\n').filter(String) };
  });
}
