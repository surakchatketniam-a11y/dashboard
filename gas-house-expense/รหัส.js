/**
 * Google Apps Script สำหรับแอพ "บันทึกค่าใช้จ่ายสร้างสำนักงาน"
 * เชื่อมต่อกับ Google Sheet: 1h83g2owxnyFiiW0s7WWiEmmbh1fquDp00uOtS4ETetU
 *
 * ความปลอดภัย / การชนกันของข้อมูล:
 *  - E11 รหัสเข้าถึง: ถ้ามี Script Property ชื่อ APP_TOKEN จะต้องส่ง token ให้ตรงทุกคำขอ
 *    (ถ้ายังไม่มี property นี้ ระบบทำงานแบบเดิมเพื่อไม่ให้แอปที่ใช้อยู่พัง)
 *    ตั้งรหัสครั้งแรกโดยเปิดตัวแก้ไข Apps Script แล้วรันฟังก์ชัน setupAppToken
 *  - E1 revision: ทุกการบันทึกได้ revision ใหม่ ถ้าแอปส่ง baseRevision มาไม่ตรงกับของในชีต
 *    (มีเครื่องอื่นบันทึกไปก่อน) จะไม่เขียนทับ แต่ตอบ conflict พร้อมข้อมูลฝั่งชีตกลับไป
 */

const SPREADSHEET_ID = '1h83g2owxnyFiiW0s7WWiEmmbh1fquDp00uOtS4ETetU';
const DATA_SHEET_NAME = '_AppData';
const CHUNK_SIZE = 40000;
const TOKEN_PROP = 'APP_TOKEN';
// โฟลเดอร์ Drive สำหรับเก็บรูปสลิป/ใบเสร็จ (E2: ย้ายรูปออกจากข้อมูลหลักที่ซิงค์กับชีต)
const SLIP_FOLDER_ID = '1kelNbtEpJCwdncnP6kqCHRZpCEJIzGhQ';
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;

function initAuthorize() {
  const ss = getSpreadsheet_();
  Logger.log('Authorized for spreadsheet: ' + ss.getName() + ' (' + ss.getId() + ')');
  return ss.getName();
}

/**
 * รันครั้งเดียวจากตัวแก้ไข Apps Script เพื่อสร้างรหัสเข้าถึงแบบสุ่ม แล้วดูรหัสใน "บันทึกการทำงาน" (Execution log)
 * รหัสถูกเก็บใน Script Properties (ไม่อยู่ในโค้ด) ถ้ามีอยู่แล้วจะแสดงรหัสเดิม ไม่สร้างใหม่
 * ต้องการเปลี่ยนรหัส: ลบ property APP_TOKEN ใน Project Settings แล้วรันใหม่ (ต้องกรอกรหัสใหม่ในแอปทุกเครื่อง)
 */
function setupAppToken() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(TOKEN_PROP);
  if (!token) {
    token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty(TOKEN_PROP, token);
    Logger.log('สร้างรหัสเข้าถึงใหม่แล้ว');
  } else {
    Logger.log('มีรหัสเข้าถึงอยู่แล้ว (แสดงรหัสเดิม)');
  }
  Logger.log('APP_TOKEN = ' + token);
  return token;
}

// ---------- รูปสลิปใน Google Drive ----------
// รันครั้งเดียวจากตัวแก้ไข Apps Script: ขออนุญาตใช้ Drive และตรวจว่าโฟลเดอร์ปลอดภัยพอจะเก็บสลิป
function checkSlipFolder() {
  const folder = getSlipFolder_();
  const access = String(folder.getSharingAccess());
  Logger.log('โฟลเดอร์: ' + folder.getName());
  Logger.log('การแชร์ของโฟลเดอร์: ' + access);
  if (access === 'PRIVATE') {
    Logger.log('OK: โฟลเดอร์เป็นแบบจำกัด พร้อมเก็บสลิป');
  } else {
    Logger.log('คำเตือน: โฟลเดอร์ถูกแชร์แบบ ' + access + ' — ระบบจะไม่อัปโหลดสลิปจนกว่าจะเปลี่ยนเป็น "จำกัด (Restricted)" เพราะสลิปมีข้อมูลบัญชีและยอดเงิน');
  }
  return access;
}

function getSlipFolder_() {
  return DriveApp.getFolderById(SLIP_FOLDER_ID);
}

function slipFolderIsPrivate_(folder) {
  return folder.getSharingAccess() === DriveApp.Access.PRIVATE;
}

function safeFileName_(name, ext) {
  let n = String(name || ('slip-' + Date.now())).replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 80);
  if (!/\.(jpg|png|webp)$/i.test(n)) n += '.' + ext;
  return n;
}

// เข้าถึงได้เฉพาะไฟล์ที่อยู่ในโฟลเดอร์สลิป (กันการอ่าน/ลบไฟล์อื่นใน Drive ด้วย fileId ที่เดามา)
function getSlipFile_(fileId) {
  if (!/^[A-Za-z0-9_-]{10,}$/.test(String(fileId || ''))) return null;
  let file;
  try { file = DriveApp.getFileById(fileId); } catch (_) { return null; }
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === SLIP_FOLDER_ID) return file;
  }
  return null;
}

function uploadImage_(payload) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(payload.dataUrl || ''));
  if (!m) return { ok: false, error: 'รูปแบบรูปไม่ถูกต้อง (รองรับ JPG/PNG/WebP)' };
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_IMAGE_BYTES) return { ok: false, error: 'รูปใหญ่เกิน 4MB' };
  const folder = getSlipFolder_();
  if (!slipFolderIsPrivate_(folder)) {
    return { ok: false, code: 'FOLDER_SHARED', error: 'โฟลเดอร์ Drive ถูกแชร์แบบลิงก์/สาธารณะ — เปลี่ยนเป็น "จำกัด (Restricted)" ก่อน ระบบจึงจะเก็บสลิป' };
  }
  const ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/webp' ? 'webp' : 'jpg');
  const file = folder.createFile(Utilities.newBlob(bytes, m[1], safeFileName_(payload.name, ext)));
  return { ok: true, fileId: file.getId(), name: file.getName() };
}

function getImage_(payload) {
  const file = getSlipFile_(payload.fileId);
  if (!file) return { ok: false, error: 'ไม่พบรูปในโฟลเดอร์สลิป' };
  const blob = file.getBlob();
  return { ok: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

function trashImage_(payload) {
  const file = getSlipFile_(payload.fileId);
  if (!file) return { ok: false, error: 'ไม่พบรูปในโฟลเดอร์สลิป' };
  file.setTrashed(true);
  return { ok: true };
}

function getExpectedToken_() {
  return PropertiesService.getScriptProperties().getProperty(TOKEN_PROP) || '';
}

function safeEqual_(a, b) {
  const x = String(a == null ? '' : a);
  const y = String(b == null ? '' : b);
  let diff = x.length ^ y.length;
  const n = Math.max(x.length, y.length);
  for (let i = 0; i < n; i++) {
    diff |= (x.charCodeAt(i) || 0) ^ (y.charCodeAt(i) || 0);
  }
  return diff === 0;
}

// ยังไม่ตั้ง APP_TOKEN = อนุญาตทุกคำขอ (โหมดเดิม) ตั้งแล้ว = ต้องส่ง token ให้ตรง
function isAuthorized_(token) {
  const expected = getExpectedToken_();
  if (!expected) return true;
  return safeEqual_(token, expected);
}

function unauthorized_() {
  return jsonResponse({ ok: false, code: 'AUTH', error: 'unauthorized: ต้องใส่รหัสเข้าถึง (Access token) ให้ถูกต้อง' });
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
}

// GET ใช้ได้เฉพาะตอนยังไม่ตั้งรหัส (แอปเวอร์ชันใหม่ใช้ POST เพื่อไม่ให้รหัสไปอยู่ใน URL)
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'load';
    if (getExpectedToken_()) {
      return jsonResponse({ ok: false, code: 'AUTH', protected: true, error: 'unauthorized: ต้องใช้รหัสเข้าถึง' });
    }
    return handleRead_(action);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function handleRead_(action) {
  const ss = getSpreadsheet_();
  if (action === 'ping') {
    return jsonResponse({
      ok: true,
      message: 'เชื่อมต่อ Google Sheets สำเร็จ',
      sheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      revision: getRevision_(),
      protected: Boolean(getExpectedToken_())
    });
  }
  if (action === 'load') {
    const data = loadAppData_();
    return jsonResponse({
      ok: true,
      hasData: data !== null,
      data: data,
      revision: getRevision_(),
      sheetName: ss.getName(),
      serverTime: new Date().toISOString()
    });
  }
  return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const action = payload.action || 'save';

    if (!isAuthorized_(payload.token)) {
      return unauthorized_();
    }

    if (action === 'ping' || action === 'load') {
      return handleRead_(action);
    }

    if (action === 'uploadImage' || action === 'getImage' || action === 'trashImage') {
      // รูปสลิปมีข้อมูลการเงิน: ต้องตั้งรหัสเข้าถึงก่อนจึงจะใช้ Drive ได้
      if (!getExpectedToken_()) {
        return jsonResponse({ ok: false, error: 'ต้องตั้งรหัสเข้าถึง (setupAppToken) ก่อนใช้ Drive' });
      }
      if (action === 'uploadImage') return jsonResponse(uploadImage_(payload));
      if (action === 'getImage') return jsonResponse(getImage_(payload));
      return jsonResponse(trashImage_(payload));
    }

    if (action === 'save') {
      if (!payload.data) {
        return jsonResponse({ ok: false, error: 'Missing data payload' });
      }

      // กันสองคำขอเขียนพร้อมกัน: ตรวจ revision กับบันทึกต้องเกิดต่อเนื่องกัน
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const current = getRevision_();
        const hasBase = Object.prototype.hasOwnProperty.call(payload, 'baseRevision');
        // แอปเวอร์ชันเก่าไม่ส่ง baseRevision = ไม่ตรวจ (เพื่อไม่ให้ของเดิมพัง)
        // แอปเวอร์ชันใหม่ส่งเสมอ (null = ยังไม่เคยรู้ revision) และไม่ตรงกับในชีต = ชนกัน
        if (hasBase && !payload.force && current && payload.baseRevision !== current) {
          return jsonResponse({
            ok: false,
            conflict: true,
            error: 'conflict: ข้อมูลในชีตถูกบันทึกจากเครื่องอื่นหลังจากที่เครื่องนี้ดึงข้อมูลมาล่าสุด',
            serverRevision: current,
            serverData: loadAppData_(),
            serverTime: new Date().toISOString()
          });
        }
        const revision = saveAppData_(payload.data);
        renderReadableSheets_(payload.data);
        return jsonResponse({
          ok: true,
          message: 'บันทึกข้อมูลลง Google Sheets เรียบร้อยแล้ว',
          revision: revision,
          sheetName: getSpreadsheet_().getName(),
          serverTime: new Date().toISOString()
        });
      } finally {
        lock.releaseLock();
      }
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// revision = รหัสรุ่นที่เก็บในเซลล์ B1 ของชีตข้อมูล (เปลี่ยนทุกครั้งที่บันทึก เวลาจริงอยู่ที่ C1) ถ้ายังไม่มีข้อมูลคืน null
function getRevision_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const v = sheet.getRange(1, 2).getValue();
  if (!v) return null;
  return (v instanceof Date) ? v.toISOString() : String(v);
}

function saveAppData_(data) {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, DATA_SHEET_NAME);
  sheet.clearContents();

  const jsonStr = JSON.stringify(data);
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    chunks.push([jsonStr.slice(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push(['{}']);

  // revision เป็นข้อความขึ้นต้นด้วยตัวอักษร เพื่อไม่ให้ Google Sheets แปลงเป็นวันที่/ตัวเลข (ซึ่งทำให้เทียบไม่ตรง)
  const now = new Date();
  const revision = 'r' + now.getTime().toString(36) + Math.random().toString(36).slice(2, 8);
  sheet.getRange(1, 2).setNumberFormat('@');
  sheet.getRange(1, 1, 1, 3).setValues([['updatedAt', revision, now.toISOString()]]);
  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);
  try { sheet.hideSheet(); } catch (_) {}
  return revision;
}

function loadAppData_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const jsonStr = values.map(function (r) { return r[0] || ''; }).join('');
  if (!jsonStr) return null;
  return JSON.parse(jsonStr);
}

function renderReadableSheets_(data) {
  const ss = getSpreadsheet_();
  const bills = data.bills || [];
  const workers = data.workers || [];
  const attendance = data.attendance || [];
  const closedPeriods = data.closedPeriods || [];
  const currentPeriodNo = data.currentPeriodNo || (closedPeriods.length + 1);
  const payments = data.payments || [];

  let ownerMatTotal = 0;
  let contractorMatTotal = 0;
  bills.forEach(function (b) {
    const bSum = (b.items || []).reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0);
    if (b.buyerType === 'owner') {
      ownerMatTotal += bSum;
    } else {
      contractorMatTotal += bSum;
    }
  });
  const totalMaterials = ownerMatTotal + contractorMatTotal;
  const paidMaterials = payments
    .filter(function (p) { return p.category === 'materials'; })
    .reduce(function (sum, p) { return sum + (Number(p.amount) || 0); }, 0);
  const totalPaidMaterials = ownerMatTotal + paidMaterials;
  const remainMaterials = contractorMatTotal - paidMaterials;

  const activeWorkerTotals = {};
  const allWorkerTotals = {};
  const allWorkerDays = {};
  workers.forEach(function (w) {
    activeWorkerTotals[w.id] = 0;
    allWorkerTotals[w.id] = 0;
    allWorkerDays[w.id] = 0;
  });

  let activeLabor = 0;
  attendance.forEach(function (day) {
    const entries = day.entries || {};
    workers.forEach(function (w) {
      const amt = Number(entries[w.id]) || 0;
      activeWorkerTotals[w.id] = (activeWorkerTotals[w.id] || 0) + amt;
      allWorkerTotals[w.id] = (allWorkerTotals[w.id] || 0) + amt;
      if (amt > 0) allWorkerDays[w.id] = (allWorkerDays[w.id] || 0) + 1;
      activeLabor += amt;
    });
  });

  let closedLabor = 0;
  closedPeriods.forEach(function (cp) {
    closedLabor += Number(cp.totalLabor) || 0;
    (cp.workerSummary || []).forEach(function (ws) {
      const matched = workers.filter(function (w) { return w.id === ws.id || w.name === ws.name; })[0];
      if (matched) {
        allWorkerTotals[matched.id] = (allWorkerTotals[matched.id] || 0) + (Number(ws.total) || 0);
        allWorkerDays[matched.id] = (allWorkerDays[matched.id] || 0) + (Number(ws.days) || 0);
      }
    });
  });

  const totalLabor = closedLabor + activeLabor;
  const paidLabor = payments
    .filter(function (p) { return p.category === 'labor'; })
    .reduce(function (sum, p) { return sum + (Number(p.amount) || 0); }, 0);
  const remainLabor = totalLabor - paidLabor;

  const activePaidLabor = payments
    .filter(function (p) { return p.category === 'labor' && Number(p.periodNo || 1) === Number(currentPeriodNo); })
    .reduce(function (sum, p) { return sum + (Number(p.amount) || 0); }, 0);
  const activeRemainLabor = activeLabor - activePaidLabor;

  // 1. สรุปภาพรวม
  const summarySheet = getOrCreateSheet_(ss, 'สรุปภาพรวม');
  summarySheet.clear();
  const projectName = data.projectName || 'โครงการสร้างบ้าน';
  const summaryRows = [
    [projectName + ' — สรุปค่าใช้จ่ายทั้งหมด', '', '', ''],
    ['อัปเดตล่าสุด', Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'), '', ''],
    ['', '', '', ''],
    ['หมวดหมู่', 'ยอดรวม (บาท)', 'จ่ายแล้ว (บาท)', 'คงเหลือต้องจ่ายช่าง (บาท)'],
    ['ค่าอุปกรณ์ (เจ้าของซื้อเอง)', ownerMatTotal, ownerMatTotal, 0],
    ['ค่าอุปกรณ์ (ช่างซื้อมาเบิก)', contractorMatTotal, paidMaterials, remainMaterials],
    ['ค่าแรงช่าง (รวมทุกงวด)', totalLabor, paidLabor, remainLabor],
    ['รวมทั้งหมด', totalMaterials + totalLabor, totalPaidMaterials + paidLabor, remainMaterials + remainLabor],
    ['', '', '', ''],
    ['สรุปค่าแรงแยกตามช่าง (สะสมทุกงวด)', 'ค่าแรง/วัน (บาท)', 'รวมค่าแรง (บาท)', 'จำนวนวันทำงาน']
  ];
  workers.forEach(function (w) {
    summaryRows.push([w.name, Number(w.dailyRate) || 0, allWorkerTotals[w.id] || 0, allWorkerDays[w.id] || 0]);
  });

  summarySheet.getRange(1, 1, summaryRows.length, 4).setValues(summaryRows);
  summarySheet.getRange(1, 1, 1, 4).merge().setFontWeight('bold').setFontSize(13).setBackground('#0a0a0a').setFontColor('#ffffff');
  summarySheet.getRange(4, 1, 1, 4).setFontWeight('bold').setBackground('#f7f7f7');
  summarySheet.getRange(8, 1, 1, 4).setFontWeight('bold').setBackground('#e6f1ed');
  summarySheet.getRange(10, 1, 1, 4).setFontWeight('bold').setBackground('#f7f7f7');
  summarySheet.getRange(5, 2, 4, 3).setNumberFormat('#,##0');
  if (workers.length > 0) {
    summarySheet.getRange(11, 2, workers.length, 2).setNumberFormat('#,##0');
  }
  summarySheet.autoResizeColumns(1, 4);

  // 2. ค่าอุปกรณ์_บิล
  const billSheet = getOrCreateSheet_(ss, 'ค่าอุปกรณ์_บิล');
  billSheet.clear();
  const billRows = [['บิล', 'ประเภทผู้ซื้อ', 'สถานะจ่ายคืนช่าง', 'วันที่', 'รายการ', 'จำนวนเงิน (บาท)', 'หมายเหตุ']];
  bills.forEach(function (b) {
    const items = b.items || [];
    const buyerLabel = b.buyerType === 'owner' ? 'เจ้าของซื้อเอง' : 'ช่างซื้อมาเบิก';
    const statusLabel = (b.buyerType === 'owner'
      ? 'เจ้าของจ่ายเองแล้ว'
      : (b.reimbursed ? ('จ่ายคืนแล้ว' + (b.reimbursedDate ? ' (' + b.reimbursedDate + ')' : '')) : 'รอจ่ายคืนช่าง')) + ((b.slipDataUrl || b.slipFileId) ? ' • มีสลิป' : '');
    let billSum = 0;
    items.forEach(function (it, idx) {
      const amt = Number(it.amount) || 0;
      billSum += amt;
      billRows.push([
        idx === 0 ? (b.name || '') : '',
        idx === 0 ? buyerLabel : '',
        idx === 0 ? statusLabel : '',
        idx === 0 ? (b.date || '') : '',
        it.name || ('รายการที่ ' + (idx + 1)),
        amt,
        idx === 0 ? (b.note || '') : ''
      ]);
    });
    billRows.push(['รวม ' + (b.name || ''), buyerLabel, statusLabel, '', '', billSum, '']);
  });
  billRows.push(['', '', '', '', '', '', '']);
  billRows.push(['รวมค่าอุปกรณ์ทั้งหมด', '', '', '', '', totalMaterials, '']);
  billRows.push(['เจ้าของซื้อเอง (จ่ายแล้ว)', '', '', '', '', ownerMatTotal, 'ไม่ค้างจ่ายช่าง']);
  billRows.push(['ช่างซื้อมาเบิก', '', '', '', '', contractorMatTotal, '']);
  billRows.push(['จ่ายคืนค่าของให้ช่างแล้ว', '', '', '', '', paidMaterials, '']);
  billRows.push(['คงเหลือต้องจ่ายช่าง', '', '', '', '', remainMaterials, '']);

  billSheet.getRange(1, 1, billRows.length, 7).setValues(billRows);
  billSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  billSheet.getRange(2, 6, billRows.length - 1, 1).setNumberFormat('#,##0');
  billSheet.getRange(billRows.length - 4, 1, 5, 7).setFontWeight('bold').setBackground('#fef3c7');
  billSheet.autoResizeColumns(1, 7);

  // 3. ค่าแรงช่าง
  const laborSheet = getOrCreateSheet_(ss, 'ค่าแรงช่าง');
  laborSheet.clear();
  const colCount = workers.length + 2;
  const headerRow = ['วันที่'].concat(workers.map(function (w) { return w.name; })).concat(['รวมรายวัน']);
  const rateRow = ['ค่าแรง/วัน'].concat(workers.map(function (w) { return Number(w.dailyRate) || 0; })).concat(['']);
  const laborRows = [headerRow, rateRow];

  attendance.forEach(function (day) {
    const entries = day.entries || {};
    let daySum = 0;
    const row = [day.date || ''];
    workers.forEach(function (w) {
      const amt = Number(entries[w.id]) || 0;
      daySum += amt;
      row.push(amt);
    });
    row.push(daySum);
    laborRows.push(row);
  });

  const totalRow = ['รวมค่าแรงงวดที่ ' + currentPeriodNo].concat(workers.map(function (w) { return activeWorkerTotals[w.id] || 0; })).concat([activeLabor]);
  const paidRow = ['เบิกล่วงหน้างวดนี้'].concat(workers.map(function () { return ''; })).concat([activePaidLabor]);
  const remainRow = ['คงเหลืองวดนี้'].concat(workers.map(function () { return ''; })).concat([activeRemainLabor]);
  laborRows.push(totalRow, paidRow, remainRow);

  laborSheet.getRange(1, 1, laborRows.length, colCount).setValues(laborRows);
  laborSheet.getRange(1, 1, 1, colCount).setFontWeight('bold').setBackground('#0a0a0a').setFontColor('#ffffff');
  laborSheet.getRange(2, 1, 1, colCount).setBackground('#f7f7f7');
  laborSheet.getRange(2, 2, laborRows.length - 1, colCount - 1).setNumberFormat('#,##0');
  laborSheet.getRange(laborRows.length - 2, 1, 3, colCount).setFontWeight('bold').setBackground('#e6f1ed');
  laborSheet.autoResizeColumns(1, colCount);

  // 4. ประวัติงวดค่าแรงที่ตัดยอดแล้ว
  const historySheet = getOrCreateSheet_(ss, 'ประวัติงวดค่าแรง');
  historySheet.clear();
  const histRows = [['งวด', 'ช่วงวันที่', 'วันที่ตัดยอด', 'จำนวนวัน', 'รวมค่าแรง (บาท)', 'จ่ายรวม (บาท)', 'สรุปรายช่าง', 'หมายเหตุ']];
  closedPeriods.forEach(function (cp) {
    const wsText = (cp.workerSummary || []).map(function (w) { return w.name + ': ' + w.total + ' (' + w.days + ' วัน)'; }).join(', ');
    histRows.push([
      cp.label || '',
      cp.dateRange || '',
      cp.closedDate || '',
      (cp.attendance || []).length,
      Number(cp.totalLabor) || 0,
      Number(cp.paidTotal) || 0,
      wsText,
      cp.note || ''
    ]);
  });
  historySheet.getRange(1, 1, histRows.length, 8).setValues(histRows);
  historySheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0a0a0a').setFontColor('#ffffff');
  if (histRows.length > 1) {
    historySheet.getRange(2, 5, histRows.length - 1, 2).setNumberFormat('#,##0');
  }
  historySheet.autoResizeColumns(1, 8);

  // 5. ประวัติการจ่ายเงิน
  const paySheet = getOrCreateSheet_(ss, 'ประวัติการจ่ายเงิน');
  paySheet.clear();
  const payRows = [['วันที่', 'หมวดหมู่', 'จำนวนเงิน (บาท)', 'สลิปโอนเงิน', 'หมายเหตุ / ผู้รับ']];
  payments.forEach(function (p) {
    payRows.push([
      p.date || '',
      p.category === 'materials' ? 'คืนค่าของ (บิลช่างซื้อ)' : 'ค่าแรงช่าง',
      Number(p.amount) || 0,
      (p.slipDataUrl || p.slipFileId) ? 'มีสลิป' : '—',
      p.note || ''
    ]);
  });
  paySheet.getRange(1, 1, payRows.length, 5).setValues(payRows);
  paySheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  if (payRows.length > 1) {
    paySheet.getRange(2, 3, payRows.length - 1, 1).setNumberFormat('#,##0');
  }
  paySheet.autoResizeColumns(1, 5);
}
