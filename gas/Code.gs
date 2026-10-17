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
        const after = rows_('bookings').find(x => x.bookingId === d.bookingId);
        const mailed = d.action === 'unconfirm' ? false : notify_(after, d.action);
        return json_({ ok: true, mailed: mailed, email: after.email || '' });
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
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(form.email || '').trim())) throw new Error('กรุณากรอกอีเมลให้ถูกต้อง เพื่อรับแจ้งผลการยืนยันนัด');
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
    if (!mine) notify_(saved, 'booked');
    return { bookingId: saved.bookingId, date: saved.date, time: saved.time, contractNo: saved.contractNo, confirmed: saved.confirmed === 'Y' };
  } finally { lock.releaseLock(); }
}

/* แจ้งบริษัททางอีเมล: booked = รับการจอง, confirm = ยืนยันนัดและเลขที่สัญญา, cancel = ยกเลิก */
function addDays_(iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function addYears_(iso, n) { const y = +iso.slice(0, 4) + n, m = iso.slice(5, 7); let dd = +iso.slice(8, 10); if (m === '02' && dd === 29 && !((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0)) dd = 28; return y + '-' + m + '-' + String(dd).padStart(2, '0'); }
function thDate_(iso) { if (!iso) return ''; const M = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม']; return (+iso.slice(8, 10)) + ' ' + M[+iso.slice(5, 7) - 1] + ' ' + (+iso.slice(0, 4) + 543); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function notify_(b, kind) {
  try {
    if (!b || !b.email) return false;
    const p = project_(b.projectId);
    const files = b.files ? JSON.parse(b.files) : [];
    const miss = p.docs.filter(d => d.required && !files.some(f => f.docKey === d.key) && (d.key !== 'poa' || b.poa === 'มอบอำนาจ') && !(d.key === 'sec' && b.secType === 'เงินสด'));
    const dur = +p.duration || 0, wy = p.warrantyYears === '' ? 2 : +p.warrantyYears, ex = p.extraDays === '' ? 30 : +p.extraDays;
    const gTo = dur ? addDays_(addYears_(addDays_(b.date, dur), wy), ex) : '';
    const link = ScriptApp.getService().getUrl() + '?p=' + encodeURIComponent(p.projectId) + '&t=' + p.token;
    const title = { booked: 'ได้รับการจองวันลงนามสัญญาแล้ว (รอเจ้าหน้าที่ยืนยัน)', confirm: 'ยืนยันวันนัดลงนามสัญญาและเลขที่สัญญาแล้ว', cancel: 'ยกเลิกการจองวันลงนามสัญญา' }[kind];
    if (!title) return false;
    const row = (k, v) => v ? '<tr><td style="padding:4px 12px 4px 0;color:#76615f;vertical-align:top;white-space:nowrap">' + k + '</td><td style="padding:4px 0">' + v + '</td></tr>' : '';
    let html = '<div style="font-family:Sarabun,Tahoma,sans-serif;font-size:15px;color:#2a1a1c;max-width:620px">'
      + '<div style="background:#7a1f2b;color:#fff;padding:14px 16px;border-radius:8px 8px 0 0"><b style="font-size:17px">' + title + '</b><br>' + esc_(p.office) + '</div>'
      + '<div style="border:1px solid #e8dcda;border-top:0;padding:14px 16px;border-radius:0 0 8px 8px"><p>เรียน ' + esc_(b.company) + ' (คุณ' + esc_(b.contactName) + ')</p>'
      + (kind === 'cancel' ? '<p>เจ้าหน้าที่ได้ยกเลิกการจองวันลงนามสัญญาของงานด้านล่าง กรุณาเปิดลิงก์เดิมเพื่อจองใหม่ หรือติดต่อเจ้าหน้าที่</p>' : '')
      + '<table style="border-collapse:collapse">' + row('งาน', esc_(p.name)) + row('ประกวดราคา', esc_(p.inviteNo))
      + (kind !== 'cancel' ? row('วันนัดลงนาม', '<b>' + thDate_(b.date) + ' เวลา ' + String(b.time).replace(':', '.') + ' น.</b>') + row('เลขที่สัญญา', '<b>' + esc_(b.contractNo) + '</b>' + (kind === 'confirm' ? '' : ' (เบื้องต้น อาจเปลี่ยนตามลำดับเวลานัด)'))
        + row('ผู้ลงนาม', b.poa === 'มอบอำนาจ' ? 'ผู้รับมอบอำนาจ (ตามหนังสือมอบอำนาจ)' : esc_(b.signerName + ' ' + b.signerSurname + ' ตำแหน่ง ' + b.signerPosition))
        + row('หลักประกันสัญญา', esc_(b.secType) + (p.secAmount ? ' จำนวน ' + Number(p.secAmount).toLocaleString('th-TH') + ' บาท' : ''))
        + row('ระยะค้ำประกัน', gTo ? 'ตั้งแต่ ' + thDate_(b.date) + ' ถึง <b>' + thDate_(gTo) + '</b>' : '') : '')
      + '</table>'
      + (kind !== 'cancel' && miss.length ? '<div style="margin-top:12px;padding:10px 12px;background:#fde6e3;color:#9d2a1f;border-radius:8px"><b>ยังไม่ได้แนบเอกสาร:</b> ' + miss.map(d => esc_(d.label)).join(', ') + '<br>กรุณาแนบผ่านลิงก์ด้านล่างก่อนวันลงนามสัญญา</div>' : '')
      + (kind === 'confirm' ? '<p>กรุณามาลงนามในสัญญาตามวันและเวลาข้างต้น' + (p.note ? ' · ' + esc_(p.note) : '') + '</p>' : '')
      + '<p><a href="' + link + '" style="display:inline-block;background:#7a1f2b;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none">เปิดหน้าการจอง / แนบเอกสาร</a></p>'
      + '<p style="color:#76615f;font-size:13px">อีเมลนี้ส่งอัตโนมัติจากระบบจองวันลงนามสัญญา ' + esc_(p.office) + ' · หากมีข้อสงสัยกรุณาติดต่อเจ้าหน้าที่</p></div></div>';
    MailApp.sendEmail({ to: b.email, subject: '[' + (p.office || 'จองลงนามสัญญา') + '] ' + title + ' · ' + String(p.name || '').slice(0, 60), htmlBody: html, name: p.office || 'ระบบจองลงนามสัญญา' });
    return true;
  } catch (e) { Logger.log('ส่งอีเมลไม่สำเร็จ: ' + e); return false; }
}

/* แนบเอกสารเพิ่มภายหลัง (เช่น หนังสือค้ำประกันจากธนาคาร) ใช้ได้แม้เจ้าหน้าที่ยืนยันแล้ว */
function attachFiles(projectId, token, newFiles) {
  project_(projectId, token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const mine = activeBookings_().find(b => b.projectId === projectId);
    if (!mine) throw new Error('ยังไม่มีการจอง');
    const old = mine.files ? JSON.parse(mine.files) : [];
    if (mine.confirmed === 'Y' && (newFiles || []).some(n => old.some(f => f.docKey === n.docKey))) throw new Error('เจ้าหน้าที่ยืนยันการจองแล้ว ไม่สามารถแก้ไขไฟล์ที่แนบไว้ได้ แนบได้เฉพาะเอกสารที่ยังขาด · หากต้องแก้ไข กรุณาติดต่อเจ้าหน้าที่');
    const files = old.filter(f => !(newFiles || []).some(n => n.docKey === f.docKey)).concat(newFiles || []);
    mine.files = JSON.stringify(files); mine.updatedAt = now_(); write_('bookings', mine);
    return files;
  } finally { lock.releaseLock(); }
}

/* Run ครั้งแรกจากหน้า Apps Script เพื่อสร้างชีตและอนุญาตสิทธิ์ */
function setup() {
  sheet_('projects'); sheet_('bookings'); sheet_('settings');
  Logger.log('โควตาอีเมลวันนี้: ' + MailApp.getRemainingDailyQuota());
  const it = DriveApp.getFoldersByName(ROOT_FOLDER); if (!it.hasNext()) DriveApp.createFolder(ROOT_FOLDER);
  Logger.log('พร้อมใช้งาน · ขั้นต่อไป: Deploy > New deployment > Web app');
}
