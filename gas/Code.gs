// ============================================================
//  TTB Mortgage Tracker — Google Apps Script (Code.gs)
//  วันที่เริ่มสัญญา: 16 พ.ค. 2026 | งวดแรก: 25 มิ.ย. 2026
//
//  วิธีอัปเดต: วางโค้ดนี้แทนของเดิมใน Apps Script แล้วไปที่
//  Deploy > Manage deployments > (ดินสอ) Edit > Version: New version > Deploy
//  (ต้องตั้ง Execute as: Me และ Who has access: Anyone)
// ============================================================

const SHEET_NAMES = {
  PAYMENTS:    'บันทึกการผ่อน',
  ACCOUNTS:    'ข้อมูลบัญชี',
  AMORTIZE:    'ตารางผ่อนชำระ',
  SUMMARY:     'สรุปภาพรวม'
};

const ACCOUNTS = [
  { id: 'HOME',  name: 'สินเชื่อตัวบ้าน',                    balance: 2153000.00 },
  { id: 'EXTRA', name: 'สินเชื่อเงินกู้เพิ่ม',               balance: 120000.00  },
  { id: 'MRTA',  name: 'ประกันชีวิตคุ้มครองสินเชื่อบ้าน (MRTA)', balance: 60232.00  }
];

const DEFAULT_RATE     = 3.0900;
const TZ               = 'Asia/Bangkok';
const CONTRACT_START   = new Date('2026-05-16T00:00:00+07:00');
const FIRST_PAYMENT    = new Date('2026-06-25T00:00:00+07:00');

// action ที่แก้ไขข้อมูล ต้องล็อกกันเขียนชนกัน และกันคำขอซ้ำด้วย reqId
const WRITE_ACTIONS = ['addPayment', 'editPayment', 'deletePayment', 'updateRate', 'init'];

// ============================================================
//  CORS Helper
// ============================================================
function corsResponse(data, callback) {
  if (callback && /^[A-Za-z_$][\w$.]*$/.test(callback)) {
    const js = callback + '(' + JSON.stringify(data) + ')';
    return ContentService
      .createTextOutput(js)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  doGet — Web App Entry Point
// ============================================================
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const cb     = params.callback || null;
  let result;

  try {
    if (WRITE_ACTIONS.indexOf(action) >= 0) {
      result = runWrite_(action, params);
    } else {
      result = route_(action, params);
    }
  } catch (err) {
    result = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }

  return corsResponse(result, cb);
}

function route_(action, params) {
  switch (action) {
    case 'getAccounts':     return getAccounts();
    case 'getPayments':     return getPayments(params);
    case 'getSummary':      return getSummary();
    case 'getAmortize':     return getAmortize(params);
    case 'addPayment':      return addPayment(params);
    case 'editPayment':     return editPayment(params);
    case 'deletePayment':   return deletePayment(params);
    case 'updateRate':      return updateRate(params);
    case 'getSettings':     return getSettings();
    case 'init':            return initSheets();
    case '':
      return { ok: true, message: 'TTB Mortgage Tracker API ready', version: '1.1' };
    default:
      // เดิมตอบ ok:true ทำให้หน้าเว็บคิดว่าบันทึกสำเร็จทั้งที่ไม่ได้ทำอะไร
      throw new Error('ไม่รู้จักคำสั่ง: ' + action + ' (อาจยังไม่ได้ Deploy เวอร์ชันใหม่)');
  }
}

// ล็อกสคริปต์ระหว่างเขียน + ถ้าเป็นคำขอซ้ำ (reqId เดิม) ให้คืนผลเดิมโดยไม่เขียนซ้ำ
function runWrite_(action, params) {
  const cache = CacheService.getScriptCache();
  const key   = params.reqId ? 'req_' + String(params.reqId).slice(0, 200) : null;

  if (key) {
    const cached = cache.get(key);
    if (cached) return Object.assign(JSON.parse(cached), { duplicate: true });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    throw new Error('ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่อีกครั้ง');
  }
  try {
    if (key) {
      // เช็คซ้ำอีกรอบหลังได้ล็อก เผื่อคำขอแรกเพิ่งเขียนเสร็จ
      const cached = cache.get(key);
      if (cached) return Object.assign(JSON.parse(cached), { duplicate: true });
    }
    const result = route_(action, params);
    SpreadsheetApp.flush();
    if (key && result && result.ok) cache.put(key, JSON.stringify(result), 21600); // 6 ชม.
    return result;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  Helpers
// ============================================================
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}

// ค่าที่ส่งมาว่างให้เป็น null (ใช้ค่าคำนวณ) แต่ "0" ต้องเป็น 0 จริง
// (ของเดิมใช้ `parseFloat(x) || ค่าอื่น` ทำให้ดอกเบี้ย 0 ถูกแทนด้วยดอกเบี้ยทั้งเดือน)
function optNum_(v) {
  if (v === undefined || v === null || String(v).trim() === '') return null;
  const n = parseFloat(String(v).replace(/,/g, ''));
  if (!isFinite(n)) throw new Error('ตัวเลขไม่ถูกต้อง: ' + v);
  return n;
}

function round2_(n) { return Math.round(n * 100) / 100; }

// "2026-09-26" -> Date เที่ยงคืนเวลาไทย
function parseDate_(s) {
  const str = String(s || '').trim();
  const m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const d = m ? new Date(str + 'T00:00:00+07:00') : new Date(str);
  if (isNaN(d.getTime())) throw new Error('วันที่ไม่ถูกต้อง: ' + s);
  return d;
}

function calcPeriod_(payDate) {
  const diffMonth = Math.round((payDate - FIRST_PAYMENT) / (1000*60*60*24*30.44)) + 1;
  return Math.max(1, diffMonth);
}

// รหัสใหม่ = เลขมากสุด + 1 (ของเดิมใช้จำนวนแถว ทำให้ ID ซ้ำหลังลบรายการ)
function nextPaymentId_(paySheet) {
  const last = paySheet.getLastRow();
  let max = 0;
  if (last > 1) {
    paySheet.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
      const m = String(r[0]).match(/^PAY(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'PAY' + String(max + 1).padStart(5, '0');
}

function findAccountRow_(accData, accountId) {
  return accData.findIndex((r, i) => i > 0 && String(r[0]).trim() === String(accountId).trim());
}

function safeRefreshSummary_() {
  // ถ้าอัปเดตหน้าสรุปพัง ไม่ควรทำให้การบันทึกทั้งหมดถูกแจ้งว่าล้มเหลว
  // (เดิม error ตรงนี้ทำให้ผู้ใช้กดบันทึกซ้ำ ข้อมูลเลยซ้ำ)
  try { refreshSummarySheet_(); } catch (err) { console.error('refreshSummarySheet_ failed: ' + err); }
}

// ============================================================
//  INIT — สร้าง Sheets เมื่อเรียกครั้งแรก
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let accSheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  if (!accSheet) {
    accSheet = ss.insertSheet(SHEET_NAMES.ACCOUNTS);
    accSheet.getRange(1,1,1,5).setValues([['รหัส','ชื่อบัญชี','ยอดเงินกู้เริ่มต้น','ยอดคงเหลือปัจจุบัน','อัตราดอกเบี้ย (%)']]);
    accSheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white');
    ACCOUNTS.forEach((a,i) => {
      accSheet.getRange(i+2,1,1,5).setValues([[a.id, a.name, a.balance, a.balance, DEFAULT_RATE]]);
    });
    accSheet.autoResizeColumns(1,5);
  }

  let paySheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!paySheet) {
    paySheet = ss.insertSheet(SHEET_NAMES.PAYMENTS);
    const headers = ['ID','วันที่ชำระ','งวดที่','รหัสบัญชี','ชื่อบัญชี','ยอดชำระทั้งหมด','เงินต้น','ดอกเบี้ย','ยอดคงเหลือหลังชำระ','หมายเหตุ','วันที่บันทึก'];
    paySheet.getRange(1,1,1,headers.length).setValues([headers]);
    paySheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white');
    paySheet.setFrozenRows(1);
    paySheet.autoResizeColumns(1,headers.length);
  }

  let amoSheet = ss.getSheetByName(SHEET_NAMES.AMORTIZE);
  if (!amoSheet) {
    amoSheet = ss.insertSheet(SHEET_NAMES.AMORTIZE);
    const hdr = ['งวดที่','วันที่ครบกำหนด','รหัสบัญชี','ชื่อบัญชี','ยอดชำระ','เงินต้น','ดอกเบี้ย','ยอดคงเหลือ'];
    amoSheet.getRange(1,1,1,hdr.length).setValues([hdr]);
    amoSheet.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white');
    amoSheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_NAMES.SUMMARY)) ss.insertSheet(SHEET_NAMES.SUMMARY);

  safeRefreshSummary_();
  return { ok: true, message: 'Sheets initialized successfully' };
}

// ============================================================
//  getAccounts
// ============================================================
function getAccounts() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  if (!sheet) throw new Error('ไม่พบชีท "' + SHEET_NAMES.ACCOUNTS + '" (เรียก action=init เพื่อสร้าง)');

  const data  = sheet.getDataRange().getValues();
  const rows  = data.slice(1)
    .filter(r => String(r[0]).trim() !== '')       // ข้ามแถวว่าง
    .map(r => ({
      id:              String(r[0]).trim(),
      name:            r[1],
      originalBalance: num_(r[2]),
      balance:         num_(r[3]),
      rate:            num_(r[4])
    }));
  return { ok: true, accounts: rows };
}

// ============================================================
//  getPayments
// ============================================================
function getPayments(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!sheet) return { ok: true, payments: [] };

  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, payments: [] };

  let rows = data.slice(1)
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      id:          String(r[0]).trim(),
      date:        r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]),
      period:      r[2],
      accountId:   String(r[3]).trim(),
      accountName: r[4],
      total:       num_(r[5]),
      principal:   num_(r[6]),
      interest:    num_(r[7]),
      balance:     num_(r[8]),
      note:        r[9] == null ? '' : String(r[9]),
      createdAt:   r[10] instanceof Date ? Utilities.formatDate(r[10], TZ, 'yyyy-MM-dd HH:mm:ss') : r[10]
    }));

  if (params && params.accountId) {
    rows = rows.filter(r => r.accountId === params.accountId);
  }

  rows.sort((a,b) => (b.date > a.date ? 1 : b.date < a.date ? -1 : (b.id > a.id ? 1 : -1)));
  return { ok: true, payments: rows };
}

// ============================================================
//  addPayment
// ============================================================
function addPayment(params) {
  if (!params.date || !params.accountId || !params.total) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน (date, accountId, total)');
  }

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  let paySheet   = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!paySheet) { initSheets(); paySheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS); }

  const accSheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  const accData  = accSheet.getDataRange().getValues();
  const accRow   = findAccountRow_(accData, params.accountId);
  if (accRow < 0) throw new Error('ไม่พบบัญชีรหัส: ' + params.accountId);

  const currentBalance = num_(accData[accRow][3]);
  const rate           = num_(accData[accRow][4]);
  const total          = optNum_(params.total);
  if (!(total > 0)) throw new Error('ยอดชำระต้องมากกว่า 0');

  let interest  = optNum_(params.interest);
  let principal = optNum_(params.principal);
  if (interest === null)  interest  = principal !== null ? total - principal : currentBalance * rate / 100 / 12;
  if (principal === null) principal = total - interest;
  interest  = round2_(interest);
  principal = round2_(principal);
  if (interest < 0 || principal < 0) throw new Error('เงินต้นและดอกเบี้ยต้องไม่ติดลบ');
  if (Math.abs(principal + interest - total) > 0.01) {
    throw new Error('เงินต้น + ดอกเบี้ย (' + round2_(principal + interest) + ') ไม่เท่ากับยอดชำระ (' + total + ')');
  }
  if (principal > currentBalance + 0.01) throw new Error('เงินต้นมากกว่ายอดคงเหลือของบัญชี');

  const newBalance = round2_(currentBalance - principal);
  const payDate    = parseDate_(params.date);
  const newId      = nextPaymentId_(paySheet);

  paySheet.appendRow([
    newId,
    payDate,
    calcPeriod_(payDate),
    params.accountId,
    accData[accRow][1],
    total,
    principal,
    interest,
    newBalance,
    params.note || '',
    new Date()
  ]);

  accSheet.getRange(accRow+1, 4).setValue(newBalance);

  safeRefreshSummary_();

  return { ok: true, message: 'บันทึกสำเร็จ', id: newId, principal, interest, newBalance };
}

// ============================================================
//  editPayment (เดิมไม่มี — หน้าเว็บเรียกแล้วได้ ok:true แต่ข้อมูลไม่เปลี่ยน)
// ============================================================
function editPayment(params) {
  if (!params.id) throw new Error('กรุณาระบุ ID การชำระ');

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!paySheet) throw new Error('ไม่พบชีทบันทึกการผ่อน');

  const data   = paySheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r,i) => i>0 && String(r[0]).trim() === String(params.id).trim());
  if (rowIdx < 0) throw new Error('ไม่พบรายการ ID: ' + params.id);
  const row = data[rowIdx];

  const oldPrincipal = num_(row[6]);
  const accountId    = String(row[3]).trim();   // ไม่อนุญาตให้ย้ายบัญชี
  const total        = params.total !== undefined ? optNum_(params.total) : num_(row[5]);
  if (!(total > 0)) throw new Error('ยอดชำระต้องมากกว่า 0');

  let interest  = optNum_(params.interest);
  let principal = optNum_(params.principal);
  if (interest === null && principal === null) { principal = oldPrincipal; interest = total - principal; }
  if (interest === null)  interest  = total - principal;
  if (principal === null) principal = total - interest;
  interest  = round2_(interest);
  principal = round2_(principal);
  if (interest < 0 || principal < 0) throw new Error('เงินต้นและดอกเบี้ยต้องไม่ติดลบ');
  if (Math.abs(principal + interest - total) > 0.01) {
    throw new Error('เงินต้น + ดอกเบี้ย (' + round2_(principal + interest) + ') ไม่เท่ากับยอดชำระ (' + total + ')');
  }

  const accSheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  const accData  = accSheet.getDataRange().getValues();
  const accRow   = findAccountRow_(accData, accountId);
  if (accRow < 0) throw new Error('ไม่พบบัญชีรหัส: ' + accountId);

  const diff        = round2_(principal - oldPrincipal);     // ตัดต้นเพิ่ม(+)/ลด(-)
  const newAccBal   = round2_(num_(accData[accRow][3]) - diff);
  if (newAccBal < -0.01) throw new Error('เงินต้นมากกว่ายอดคงเหลือของบัญชี');
  const payDate     = params.date ? parseDate_(params.date) : row[1];
  const note        = params.note !== undefined ? params.note : row[9];

  // คอลัมน์ B..J : วันที่, งวด, รหัส, ชื่อ, ยอด, ต้น, ดอก, คงเหลือหลังชำระ, หมายเหตุ
  paySheet.getRange(rowIdx+1, 2, 1, 9).setValues([[
    payDate,
    calcPeriod_(payDate instanceof Date ? payDate : parseDate_(payDate)),
    accountId,
    row[4],
    total,
    principal,
    interest,
    round2_(num_(row[8]) - diff),
    note
  ]]);
  accSheet.getRange(accRow+1, 4).setValue(newAccBal);

  safeRefreshSummary_();
  return { ok: true, message: 'แก้ไขสำเร็จ', id: params.id, principal, interest, newBalance: newAccBal };
}

// ============================================================
//  deletePayment
// ============================================================
function deletePayment(params) {
  if (!params.id) throw new Error('กรุณาระบุ ID การชำระ');

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!paySheet) throw new Error('ไม่พบชีทบันทึกการผ่อน');

  const data = paySheet.getDataRange().getValues();
  const rowIdx = data.findIndex((r,i) => i>0 && String(r[0]).trim() === String(params.id).trim());
  if (rowIdx < 0) throw new Error('ไม่พบรายการ ID: ' + params.id);

  // คืนยอดเงินต้นที่ตัดไปกลับ
  const principal  = num_(data[rowIdx][6]);
  const accountId  = String(data[rowIdx][3]).trim();
  const accSheet   = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  const accData    = accSheet.getDataRange().getValues();
  const accRow     = findAccountRow_(accData, accountId);
  if (accRow >= 0) {
    accSheet.getRange(accRow+1, 4).setValue(round2_(num_(accData[accRow][3]) + principal));
  }

  paySheet.deleteRow(rowIdx+1);
  safeRefreshSummary_();
  return { ok: true, message: 'ลบรายการสำเร็จ' };
}

// ============================================================
//  updateRate
// ============================================================
function updateRate(params) {
  const rate = optNum_(params.rate);
  if (!params.accountId || rate === null) throw new Error('ต้องระบุ accountId และ rate');
  if (rate < 0 || rate > 30) throw new Error('อัตราดอกเบี้ยไม่สมเหตุสมผล: ' + rate);

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  const data     = accSheet.getDataRange().getValues();
  const rowIdx   = findAccountRow_(data, params.accountId);
  if (rowIdx < 0) throw new Error('ไม่พบบัญชี');

  accSheet.getRange(rowIdx+1, 5).setValue(rate);
  return { ok: true, message: 'อัปเดตดอกเบี้ยสำเร็จ' };
}

// ============================================================
//  getSettings
// ============================================================
function getSettings() {
  return {
    ok: true,
    contractStart: Utilities.formatDate(CONTRACT_START, TZ, 'yyyy-MM-dd'),
    firstPayment:  Utilities.formatDate(FIRST_PAYMENT,  TZ, 'yyyy-MM-dd'),
    defaultRate:   DEFAULT_RATE,
    bank:          'TTB',
    version:       '1.1'
  };
}

// ============================================================
//  getAmortize — ตารางผ่อนชำระแบบแสดงเดือน
// ============================================================
function getAmortize(params) {
  const accRes = getAccounts();
  const result = [];
  const months = Math.max(1, parseInt(params && params.months ? params.months : 240, 10) || 240);

  accRes.accounts.forEach(acc => {
    if (params && params.accountId && acc.id !== params.accountId) return;

    const monthlyRate = acc.rate / 100 / 12;
    const payment     = monthlyRate > 0
      ? (acc.balance * monthlyRate) / (1 - Math.pow(1 + monthlyRate, -months))
      : acc.balance / months;

    let balance = acc.balance;
    for (let m = 1; m <= months && balance > 0; m++) {
      const interest   = round2_(balance * monthlyRate);
      const principal  = round2_(Math.min(payment - interest, balance));
      balance          = round2_(balance - principal);

      const dueDate = new Date(FIRST_PAYMENT);
      dueDate.setMonth(dueDate.getMonth() + m - 1);

      result.push({
        accountId:   acc.id,
        accountName: acc.name,
        period:      m,
        dueDate:     Utilities.formatDate(dueDate, TZ, 'yyyy-MM-dd'),
        payment:     round2_(payment),
        principal,
        interest,
        balance:     Math.max(0, balance)
      });
    }
  });

  return { ok: true, rows: result };
}

// ============================================================
//  getSummary
// ============================================================
function getSummary() {
  const accRes  = getAccounts();
  const payRes  = getPayments({});

  // ใช้ num_ ทุกจุด: ถ้ามีช่องว่าง/ข้อความในชีต ของเดิมจะกลายเป็นการต่อ string แล้ว toFixed พัง
  const totalOriginal  = accRes.accounts.reduce((s,a) => s + a.originalBalance, 0);
  const totalRemaining = accRes.accounts.reduce((s,a) => s + a.balance, 0);
  const totalPaid      = payRes.payments.reduce((s,p) => s + p.total, 0);
  const totalInterest  = payRes.payments.reduce((s,p) => s + p.interest, 0);
  const totalPrincipal = payRes.payments.reduce((s,p) => s + p.principal, 0);

  return {
    ok: true,
    summary: {
      totalOriginal:  round2_(totalOriginal),
      totalRemaining: round2_(totalRemaining),
      totalPaid:      round2_(totalPaid),
      totalInterest:  round2_(totalInterest),
      totalPrincipal: round2_(totalPrincipal),
      progressPercent: totalOriginal > 0 ? round2_((totalOriginal - totalRemaining) / totalOriginal * 100) : 0,
      paymentCount:   payRes.payments.length
    },
    accounts: accRes.accounts
  };
}

// ============================================================
//  refreshSummarySheet_ (private)
// ============================================================
function refreshSummarySheet_() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  let sumSheet   = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!sumSheet) sumSheet = ss.insertSheet(SHEET_NAMES.SUMMARY);

  const summary  = getSummary();
  const s        = summary.summary;

  const rows = [
    ['สรุปภาพรวมการผ่อนบ้าน', '', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')],
    ['', '', ''],
    ['รายการ', 'จำนวนเงิน (บาท)', ''],
    ['ยอดกู้รวมทั้งหมด',  s.totalOriginal,  ''],
    ['ยอดคงเหลือรวม',     s.totalRemaining, ''],
    ['ชำระแล้วรวม',       s.totalPaid,      ''],
    ['  - เงินต้น',       s.totalPrincipal, ''],
    ['  - ดอกเบี้ย',      s.totalInterest,  ''],
    ['จำนวนรายการที่ชำระ', s.paymentCount,   'รายการ'],
    ['ความคืบหน้า',        s.progressPercent,'%'],
    ['', '', ''],
    ['รายละเอียดบัญชี','ยอดเริ่มต้น','ยอดคงเหลือ']
  ];

  summary.accounts.forEach(a => {
    rows.push([a.name, a.originalBalance, a.balance]);
  });

  sumSheet.clearContents();
  sumSheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sumSheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1B4F72').setFontColor('white').setFontSize(12);
  sumSheet.getRange(3,1,1,3).setFontWeight('bold').setBackground('#AED6F1');
  sumSheet.getRange(12,1,1,3).setFontWeight('bold').setBackground('#AED6F1');

  const numFmt = '#,##0.00';
  sumSheet.getRange(4, 2, 5, 1).setNumberFormat(numFmt);
  sumSheet.getRange(9, 2, 1, 1).setNumberFormat('0');
  sumSheet.getRange(10, 2, 1, 1).setNumberFormat('0.00');
  if (summary.accounts.length) {
    sumSheet.getRange(13, 2, summary.accounts.length, 2).setNumberFormat(numFmt);
  }
}
