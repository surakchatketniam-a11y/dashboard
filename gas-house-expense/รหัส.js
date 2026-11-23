/**
 * Google Apps Script สำหรับแอพ "บันทึกค่าใช้จ่ายสร้างบ้าน"
 * เชื่อมต่อกับ Google Sheet: 1h83g2owxnyFiiW0s7WWiEmmbh1fquDp00uOtS4ETetU
 */

const SPREADSHEET_ID = '1h83g2owxnyFiiW0s7WWiEmmbh1fquDp00uOtS4ETetU';
const DATA_SHEET_NAME = '_AppData';
const CHUNK_SIZE = 40000;

function initAuthorize() {
  const ss = getSpreadsheet_();
  Logger.log('Authorized for spreadsheet: ' + ss.getName() + ' (' + ss.getId() + ')');
  return ss.getName();
}

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'load';
    const ss = getSpreadsheet_();
    if (action === 'ping') {
      return jsonResponse({
        ok: true,
        message: 'เชื่อมต่อ Google Sheets สำเร็จ',
        sheetName: ss.getName(),
        spreadsheetId: ss.getId()
      });
    }
    if (action === 'load') {
      const data = loadAppData_();
      return jsonResponse({
        ok: true,
        hasData: data !== null,
        data: data,
        sheetName: ss.getName(),
        serverTime: new Date().toISOString()
      });
    }
    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(raw);
    const action = payload.action || 'save';

    if (action === 'save') {
      if (!payload.data) {
        return jsonResponse({ ok: false, error: 'Missing data payload' });
      }
      saveAppData_(payload.data);
      renderReadableSheets_(payload.data);
      return jsonResponse({
        ok: true,
        message: 'บันทึกข้อมูลลง Google Sheets เรียบร้อยแล้ว',
        sheetName: getSpreadsheet_().getName(),
        serverTime: new Date().toISOString()
      });
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

  sheet.getRange(1, 1, 1, 2).setValues([['updatedAt', new Date().toISOString()]]);
  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);
  try { sheet.hideSheet(); } catch (_) {}
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

  const totalMaterials = bills.reduce(function (sum, b) {
    return sum + (b.items || []).reduce(function (s, it) { return s + (Number(it.amount) || 0); }, 0);
  }, 0);
  const paidMaterials = payments
    .filter(function (p) { return p.category === 'materials'; })
    .reduce(function (sum, p) { return sum + (Number(p.amount) || 0); }, 0);
  const remainMaterials = totalMaterials - paidMaterials;

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
    ['หมวดหมู่', 'ยอดรวม (บาท)', 'จ่ายแล้ว (บาท)', 'คงเหลือ (บาท)'],
    ['ซื้ออุปกรณ์ / บิล', totalMaterials, paidMaterials, remainMaterials],
    ['ค่าแรงช่าง (รวมทุกงวด)', totalLabor, paidLabor, remainLabor],
    ['รวมทั้งหมด', totalMaterials + totalLabor, paidMaterials + paidLabor, remainMaterials + remainLabor],
    ['', '', '', ''],
    ['สรุปค่าแรงแยกตามช่าง (สะสมทุกงวด)', 'ค่าแรง/วัน (บาท)', 'รวมค่าแรง (บาท)', 'จำนวนวันทำงาน']
  ];
  workers.forEach(function (w) {
    summaryRows.push([w.name, Number(w.dailyRate) || 0, allWorkerTotals[w.id] || 0, allWorkerDays[w.id] || 0]);
  });

  summarySheet.getRange(1, 1, summaryRows.length, 4).setValues(summaryRows);
  summarySheet.getRange(1, 1, 1, 4).merge().setFontWeight('bold').setFontSize(13).setBackground('#0a0a0a').setFontColor('#ffffff');
  summarySheet.getRange(4, 1, 1, 4).setFontWeight('bold').setBackground('#f7f7f7');
  summarySheet.getRange(7, 1, 1, 4).setFontWeight('bold').setBackground('#e6f1ed');
  summarySheet.getRange(9, 1, 1, 4).setFontWeight('bold').setBackground('#f7f7f7');
  summarySheet.getRange(5, 2, 3, 3).setNumberFormat('#,##0');
  if (workers.length > 0) {
    summarySheet.getRange(10, 2, workers.length, 2).setNumberFormat('#,##0');
  }
  summarySheet.autoResizeColumns(1, 4);

  // 2. ค่าอุปกรณ์_บิล
  const billSheet = getOrCreateSheet_(ss, 'ค่าอุปกรณ์_บิล');
  billSheet.clear();
  const billRows = [['บิล', 'วันที่', 'รายการ', 'จำนวนเงิน (บาท)', 'หมายเหตุ']];
  bills.forEach(function (b) {
    const items = b.items || [];
    let billSum = 0;
    items.forEach(function (it, idx) {
      const amt = Number(it.amount) || 0;
      billSum += amt;
      billRows.push([
        idx === 0 ? (b.name || '') : '',
        idx === 0 ? (b.date || '') : '',
        it.name || ('รายการที่ ' + (idx + 1)),
        amt,
        idx === 0 ? (b.note || '') : ''
      ]);
    });
    billRows.push(['รวม ' + (b.name || ''), '', '', billSum, '']);
  });
  billRows.push(['', '', '', '', '']);
  billRows.push(['รวมค่าอุปกรณ์ทั้งหมด', '', '', totalMaterials, '']);
  billRows.push(['จ่ายแล้ว', '', '', paidMaterials, '']);
  billRows.push(['คงเหลือ', '', '', remainMaterials, '']);

  billSheet.getRange(1, 1, billRows.length, 5).setValues(billRows);
  billSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  billSheet.getRange(2, 4, billRows.length - 1, 1).setNumberFormat('#,##0');
  billSheet.getRange(billRows.length - 2, 1, 3, 5).setFontWeight('bold').setBackground('#fef3c7');
  billSheet.autoResizeColumns(1, 5);

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

  // 4. ประวัติการจ่ายเงิน
  const paySheet = getOrCreateSheet_(ss, 'ประวัติการจ่ายเงิน');
  paySheet.clear();
  const payRows = [['วันที่', 'หมวดหมู่', 'จำนวนเงิน (บาท)', 'หมายเหตุ / ผู้รับ']];
  payments.forEach(function (p) {
    payRows.push([
      p.date || '',
      p.category === 'materials' ? 'ค่าอุปกรณ์ / บิล' : 'ค่าแรงช่าง',
      Number(p.amount) || 0,
      p.note || ''
    ]);
  });
  paySheet.getRange(1, 1, payRows.length, 4).setValues(payRows);
  paySheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  if (payRows.length > 1) {
    paySheet.getRange(2, 3, payRows.length - 1, 1).setNumberFormat('#,##0');
  }
  paySheet.autoResizeColumns(1, 4);
}
