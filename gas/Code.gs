/**
 * ระบบจองวันลงนามสัญญาออนไลน์ (ใช้คู่กับ "ระบบพิจารณาผล - ต่อรอง/ปรับลดราคา - เรียกลงนามสัญญา")
 * วิธีติดตั้ง: ดูไฟล์ README ในโฟลเดอร์ gas หรือคู่มือในแอป หน้าตั้งค่า
 *
 * ขั้นตอนเดียวที่ต้องแก้ในไฟล์นี้: เปลี่ยน ADMIN_KEY ด้านล่างเป็นรหัสลับของหน่วยงาน
 * (ต้องตรงกับช่อง "รหัสเจ้าหน้าที่" ในหน้าตั้งค่าของแอป)
 */
const ADMIN_KEY = 'เปลี่ยนเป็นรหัสลับของหน่วยงาน';

const TZ = 'Asia/Bangkok';
const ROOT_FOLDER = 'เอกสารจองลงนามสัญญา';
const MAX_FILE_MB = 10;
/* หลักประกันสัญญา ตามระเบียบกระทรวงการคลังว่าด้วยการจัดซื้อจัดจ้างและการบริหารพัสดุภาครัฐ พ.ศ. 2560 ข้อ 167 */
/* หนังสือค้ำประกันต้องขอจากธนาคาร/บริษัทเงินทุนก่อน จึงแนบภายหลังได้ */
const SEC_LATER = ['หนังสือค้ำประกันของธนาคารภายในประเทศ', 'หนังสือค้ำประกันอิเล็กทรอนิกส์ (e-Guarantee)', 'หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์'];
const SEC_TYPES = ['เงินสด', 'เช็คหรือดราฟท์ที่ธนาคารเซ็นสั่งจ่าย', 'หนังสือค้ำประกันของธนาคารภายในประเทศ', 'หนังสือค้ำประกันอิเล็กทรอนิกส์ (e-Guarantee)', 'หนังสือค้ำประกันของบริษัทเงินทุนหรือบริษัทเงินทุนหลักทรัพย์', 'พันธบัตรรัฐบาลไทย'];
const COLS = {
  projects: ['projectId', 'token', 'name', 'contractor', 'inviteNo', 'office', 'dates', 'slotStart', 'slotEnd', 'slotMinutes', 'docs', 'note', 'updatedAt', 'secAmount', 'stamp', 'duration', 'warrantyYears', 'extraDays', 'beneficiary'],
  bookings: ['bookingId', 'projectId', 'date', 'time', 'company', 'contactName', 'phone', 'email', 'poa', 'signerName', 'signerSurname', 'signerPosition', 'files', 'status', 'contractNo', 'confirmed', 'createdAt', 'updatedAt', 'seq', 'secType'],
  settings: ['key', 'value']
};

/* ---------------- sheets ---------------- */
function book_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active) throw new Error('เปิด Apps Script จากเมนู ส่วนขยาย ของ Google Sheet แล้วกด Run ฟังก์ชัน setup ก่อน');
    id = active.getId();
    props.setProperty('SHEET_ID', id);
  }
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1000, 30).setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    const n = COLS[name].length, h = sh.getRange(1, 1, 1, n).getValues()[0];
    if (h.join('|') !== COLS[name].join('|')) sh.getRange(1, 1, 1, n).setValues([COLS[name]]).setFontWeight('bold');
  }
  return sh;
}
function norm_(v, k) {
  if (v instanceof Date) {
    if (k === 'time') return Utilities.formatDate(v, TZ, 'HH:mm');
    if (k === 'date') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
  }
  return v === null || v === undefined ? '' : String(v);
}
function rows_(name) {
  const sh = sheet_(name);
  const v = sh.getDataRange().getValues();
  const h = v.shift() || [];
  return v.filter(r => r.some(c => c !== '')).map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => o[k] = norm_(r[j], k));
    return o;
  });
}
function write_(name, obj) {
  const sh = sheet_(name);
  const row = COLS[name].map(k => obj[k] === undefined ? '' : String(obj[k]));
  if (obj._row) sh.getRange(obj._row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}
function settings_() {
  const o = { contractFormat: '{n}/{yyyy}', contractStart: '1', contractFY: String(fiscalYear_(today_())) };
  rows_('settings').forEach(r => { if (r.key) o[r.key] = r.value; });
  return o;
}
function setSetting_(key, value) {
  const r = rows_('settings').find(x => x.key === key);
  write_('settings', { _row: r ? r._row : undefined, key: key, value: value });
}

/* ---------------- helpers ---------------- */
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function fiscalYear_(iso) { const y = +iso.slice(0, 4), m = +iso.slice(5, 7); return y + 543 + (m >= 10 ? 1 : 0); }
function json_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb && /^[\w.$]+$/.test(cb)) return ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function times_(p) {
  const toMin = t => { const m = String(t || '').match(/(\d{1,2})[.:](\d{2})/); return m ? (+m[1]) * 60 + (+m[2]) : null; };
  const a = toMin(p.slotStart) ?? 540, b = toMin(p.slotEnd) ?? 930, step = +p.slotMinutes || 30;
  const out = [];
  for (let t = a; t <= b; t += step) out.push(String(Math.floor(t / 60)).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0'));
  return out;
}
function activeBookings_() { return rows_('bookings').filter(b => b.status !== 'ยกเลิก'); }
function project_(projectId, token) {
  const p = rows_('projects').find(x => x.projectId === projectId);
  if (!p || (token !== undefined && p.token !== token)) throw new Error('ลิงก์จองไม่ถูกต้องหรือหมดอายุ กรุณาติดต่อเจ้าหน้าที่');
  p.dates = p.dates ? JSON.parse(p.dates) : [];
  p.docs = p.docs ? JSON.parse(p.docs) : [];
  return p;
}

function formatNo_(f, n, fy) {
  return String(f || '{n}/{yyyy}').replace(/\{(n+)\}/g, (m, k) => String(n).padStart(k.length, '0')).replace(/\{yyyy\}/g, fy).replace(/\{yy\}/g, String(fy).slice(-2));
}
/* เลขที่สัญญา: เรียงตามวันเวลาที่นัดลงนาม เลขที่ยืนยันแล้วไม่เปลี่ยน */
function renumber_() {
  const st = settings_();
  const list = activeBookings_();
  const byFY = {};
  list.forEach(b => { if (!b.date) return; const fy = fiscalYear_(b.date); (byFY[fy] = byFY[fy] || []).push(b); });
  Object.keys(byFY).forEach(fy => {
    const arr = byFY[fy].sort((x, y) => (x.date + x.time).localeCompare(y.date + y.time) || x.createdAt.localeCompare(y.createdAt));
    const used = new Set(arr.filter(b => b.confirmed === 'Y' && b.seq !== '').map(b => +b.seq));
    let n = String(fy) === String(st.contractFY) ? (+st.contractStart || 1) : 1;
    arr.forEach(b => {
      if (b.confirmed === 'Y') return;
      while (used.has(n)) n++;
      const no = formatNo_(st.contractFormat, n, fy);
      if (b.contractNo !== no || String(b.seq) !== String(n)) { b.contractNo = no; b.seq = n; write_('bookings', b); }
      n++;
    });
  });
}

/* ---------------- web endpoints ---------------- */
function doGet(e) {
  const q = e.parameter || {};
  if (q.action) {
    try {
      if (q.key !== ADMIN_KEY) throw new Error('รหัสเจ้าหน้าที่ไม่ถูกต้อง');
      if (q.action === 'list') {
        const b = rows_('bookings').filter(x => !q.p || x.projectId === q.p).map(x => { delete x._row; x.files = x.files ? JSON.parse(x.files) : []; return x; });
        return json_({ ok: true, bookings: b, settings: settings_() }, q.callback);
      }
      if (q.action === 'ping') return json_({ ok: true, settings: settings_() }, q.callback);
      throw new Error('ไม่รู้จักคำสั่ง');
    } catch (err) { return json_({ ok: false, error: String(err.message || err) }, q.callback); }
  }
  if (q.p && q.t) {
    const t = HtmlService.createTemplateFromFile('Booking');
    t.projectId = q.p; t.token = q.t;
    return t.evaluate().setTitle('จองวันลงนามสัญญา').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">ระบบจองวันลงนามสัญญา · กรุณาสแกน QR code จากหนังสือเรียกลงนามสัญญา</p>');
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.key !== ADMIN_KEY) throw new Error('รหัสเจ้าหน้าที่ไม่ถูกต้อง');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      if (d.action === 'upsertProject') {
        const p = d.project;
        const old = rows_('projects').find(x => x.projectId === p.projectId);
        write_('projects', {
          _row: old ? old._row : undefined, projectId: p.projectId, token: p.token || (old ? old.token : uid_()), name: p.name, contractor: p.contractor,
          inviteNo: p.inviteNo, office: p.office, dates: JSON.stringify(p.dates || []), slotStart: p.slotStart, slotEnd: p.slotEnd,
          slotMinutes: p.slotMinutes, docs: JSON.stringify(p.docs || []), note: p.note || '', updatedAt: now_(), secAmount: p.secAmount || '', stamp: p.stamp || '', duration: p.duration || '', warrantyYears: p.warrantyYears || '', extraDays: p.extraDays === undefined ? '' : p.extraDays, beneficiary: p.beneficiary || ''
        });
        const saved = rows_('projects').find(x => x.projectId === p.projectId);
        return json_({ ok: true, token: saved.token });
      }
      if (d.action === 'settings') {
        ['contractFormat', 'contractStart', 'contractFY'].forEach(k => { if (d[k] !== undefined) setSetting_(k, d[k]); });
        renumber_();
        return json_({ ok: true, settings: settings_() });
      }
      if (d.action === 'confirm' || d.action === 'unconfirm' || d.action === 'cancel') {
        const b = rows_('bookings').find(x => x.bookingId === d.bookingId);
        if (!b) throw new Error('ไม่พบรายการจอง');
        if (d.action === 'confirm') { b.confirmed = 'Y'; b.status = 'ยืนยันแล้ว'; if (d.contractNo) b.contractNo = d.contractNo; }
        if (d.action === 'unconfirm') { b.confirmed = ''; b.status = 'จองแล้ว'; }
        if (d.action === 'cancel') { b.status = 'ยกเลิก'; b.confirmed = ''; }
        b.updatedAt = now_(); write_('bookings', b); renumber_();
        return json_({ ok: true });
      }
      throw new Error('ไม่รู้จักคำสั่ง');
    } finally { lock.releaseLock(); }
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

/* ---------------- called from Booking.html ---------------- */
function getBookingInfo(projectId, token) {
  const p = project_(projectId, token);
  const all = activeBookings_();
  const mine = all.find(b => b.projectId === projectId) || null;
  const taken = {};
  all.forEach(b => { if (!mine || b.bookingId !== mine.bookingId) taken[b.date + ' ' + b.time] = true; });
  const t = today_(), nowHM = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  const days = p.dates.filter(d => d >= t).map(d => ({ date: d, times: times_(p).map(tm => ({ time: tm, free: !taken[d + ' ' + tm] && !(d === t && tm <= nowHM) })) }));
  if (mine) mine.files = mine.files ? JSON.parse(mine.files) : [];
  return { project: { name: p.name, contractor: p.contractor, inviteNo: p.inviteNo, office: p.office, docs: p.docs, note: p.note, secAmount: p.secAmount, stamp: p.stamp, duration: p.duration, warrantyYears: p.warrantyYears, extraDays: p.extraDays, beneficiary: p.beneficiary, secTypes: SEC_TYPES, secLater: SEC_LATER }, days: days, booking: mine, maxMB: MAX_FILE_MB };
}

function uploadFile(projectId, token, docKey, fileName, mime, b64) {
  const p = project_(projectId, token);
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > MAX_FILE_MB * 1024 * 1024) throw new Error('ไฟล์ใหญ่เกิน ' + MAX_FILE_MB + ' MB');
  const it = DriveApp.getFoldersByName(ROOT_FOLDER);
  const root = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER);
  const fname = (p.inviteNo || projectId) + ' ' + String(p.name || '').slice(0, 60);
  const it2 = root.getFoldersByName(fname);
  const folder = it2.hasNext() ? it2.next() : root.createFolder(fname);
  const f = folder.createFile(Utilities.newBlob(bytes, mime || 'application/octet-stream', docKey + ' - ' + fileName));
  return { docKey: docKey, name: fileName, id: f.getId(), url: f.getUrl() };
}

function submitBooking(projectId, token, form) {
  const p = project_(projectId, token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!p.dates.includes(form.date)) throw new Error('วันที่นี้ไม่อยู่ในช่วงที่จองได้');
    if (!times_(p).includes(form.time)) throw new Error('เวลานี้จองไม่ได้');
    if (form.date < today_()) throw new Error('วันที่ผ่านมาแล้ว');
    if (!form.contactName || !form.phone) throw new Error('กรุณากรอกชื่อผู้ติดต่อและเบอร์โทร');
    if (form.poa !== 'yes' && (!form.signerName || !form.signerSurname || !form.signerPosition)) throw new Error('ไม่มอบอำนาจ กรุณาระบุชื่อ นามสกุล และตำแหน่งผู้ลงนามในสัญญา');
    const all = activeBookings_();
    let mine = all.find(b => b.projectId === projectId);
    if (mine && mine.confirmed === 'Y' && (mine.date !== form.date || mine.time !== form.time)) throw new Error('เจ้าหน้าที่ยืนยันวันนัดและเลขที่สัญญาแล้ว ถ้าต้องการเปลี่ยนวัน กรุณาติดต่อเจ้าหน้าที่');
    if (all.some(b => b.date === form.date && b.time === form.time && (!mine || b.bookingId !== mine.bookingId))) throw new Error('ช่วงเวลานี้มีผู้จองแล้ว กรุณาเลือกเวลาอื่น');
    const files = (mine && mine.files ? JSON.parse(mine.files) : []).filter(f => !(form.files || []).some(n => n.docKey === f.docKey)).concat(form.files || []);
    if (SEC_TYPES.indexOf(form.secType) < 0) throw new Error('กรุณาเลือกประเภทหลักประกันสัญญา');
    const required = p.docs.filter(d => d.required && (d.key !== 'poa' || form.poa === 'yes') && (d.key !== 'sec' || (form.secType !== 'เงินสด' && SEC_LATER.indexOf(form.secType) < 0)));
    const missing = required.filter(d => !files.some(f => f.docKey === d.key));
    if (missing.length) throw new Error('ยังไม่ได้แนบ: ' + missing.map(d => d.label).join(', '));
    const b = Object.assign(mine || { bookingId: uid_(), projectId: projectId, createdAt: now_(), status: 'จองแล้ว', confirmed: '' }, {
      date: form.date, time: form.time, company: p.contractor, contactName: form.contactName, phone: form.phone, email: form.email || '',
      poa: form.poa === 'yes' ? 'มอบอำนาจ' : 'ไม่มอบอำนาจ', signerName: form.signerName || '', signerSurname: form.signerSurname || '',
      signerPosition: form.signerPosition || '', secType: form.secType, files: JSON.stringify(files), updatedAt: now_()
    });
    write_('bookings', b);
    renumber_();
    const saved = rows_('bookings').find(x => x.bookingId === b.bookingId);
    return { bookingId: saved.bookingId, date: saved.date, time: saved.time, contractNo: saved.contractNo, confirmed: saved.confirmed === 'Y' };
  } finally { lock.releaseLock(); }
}

/* แนบเอกสารเพิ่มภายหลัง (เช่น หนังสือค้ำประกันจากธนาคาร) ใช้ได้แม้เจ้าหน้าที่ยืนยันแล้ว */
function attachFiles(projectId, token, newFiles) {
  project_(projectId, token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const mine = activeBookings_().find(b => b.projectId === projectId);
    if (!mine) throw new Error('ยังไม่มีการจอง');
    const files = (mine.files ? JSON.parse(mine.files) : []).filter(f => !(newFiles || []).some(n => n.docKey === f.docKey)).concat(newFiles || []);
    mine.files = JSON.stringify(files); mine.updatedAt = now_(); write_('bookings', mine);
    return files;
  } finally { lock.releaseLock(); }
}

/* Run ครั้งแรกจากหน้า Apps Script เพื่อสร้างชีตและอนุญาตสิทธิ์ */
function setup() {
  sheet_('projects'); sheet_('bookings'); sheet_('settings');
  const it = DriveApp.getFoldersByName(ROOT_FOLDER); if (!it.hasNext()) DriveApp.createFolder(ROOT_FOLDER);
  Logger.log('พร้อมใช้งาน · ขั้นต่อไป: Deploy > New deployment > Web app');
}
