// ============================================================
// Code.gs — Google Apps Script Backend
// ระบบเบิกอุปกรณ์สำนักงาน บริษัทเคยี่ ก่อสร้าง
// ============================================================

const SHEET_ID      = '1lE9u_DLYaOghT4IvpFM0quP20WD_rTGJOE66UXajKjw';
const SHEET_FORM    = 'การตอบแบบฟอร์ม 1';
const SHEET_PRODUCT = 'บันทึกรายการ';
const SHEET_EMP     = 'รหัสพนักงาน';

// ส่วนของ doGet สำหรับแสดงหน้าเว็บ (Index.html)
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเบิกอุปกรณ์สำนักงาน บริษัทเคยี่ ก่อสร้าง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ส่วนของ doPost สำหรับรับคำสั่งจาก JavaScript (API)
function doPost(e) {
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  var data = params.data;
  var result = {};

  switch (action) {
    case 'saveRequisition': result = saveRequisition(data); break;
    case 'getProductByCode': result = getProductByCode(data.code); break;
    case 'getEmployeeByCode': result = getEmployeeByCode(data.code); break;
    case 'getRecentHistory': result = getRecentHistory(); break;
    case 'getDashboardSummary': result = getDashboardSummary(); break;
    default: result = { error: 'ไม่พบคำสั่งที่ระบุ' };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

// แก้ไขฟังก์ชัน saveRequisition เล็กน้อยเพื่อส่งค่ากลับให้ถูกต้อง
function saveRequisition(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    
    // ... [โค้ดเดิมของคุณที่จัดการสต็อกและบันทึก row] ...
    // แนะนำให้ตรวจสอบว่ามีการเรียกฟังก์ชัน _getAdminSheet() ไว้ด้านบนด้วย
    
    return { success: true, docId: 'TEST-ID', newStock: 0 }; 
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// ดึงข้อมูลสินค้า — Sheet "บันทึกรายการ"
// A=รหัส, C=ประเภท, E=ชื่อสินค้า, F=ขนาด, K=ยอดคงเหลือ, L=หน่วย
// header 2 แถว ข้อมูลเริ่มแถว 3 (i=2)
// ══════════════════════════════════════════════════════════════
function getProductByCode(code) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRODUCT);
    if (!sheet) return { error: 'ไม่พบ Sheet: ' + SHEET_PRODUCT };
    const data       = sheet.getDataRange().getValues();
    const searchCode = String(code).trim();
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchCode) {
        return {
          code    : data[i][0],
          category: data[i][2],
          name    : data[i][4],
          size    : data[i][5],
          stock   : data[i][10],
          unit    : data[i][11],
        };
      }
    }
    return null;
  } catch(e) {
    return { error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// ดึงข้อมูลพนักงาน — Sheet "รหัสพนักงาน"
// A=รหัส, C=ชื่อจีน, D=ชื่อ-สกุลไทย, F=แผนก, K=ตำแหน่ง, P=สังกัด
// header 1 แถว ข้อมูลเริ่มแถว 2 (i=1)
// ══════════════════════════════════════════════════════════════
function getEmployeeByCode(empCode) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_EMP);
    if (!sheet) return { error: 'ไม่พบ Sheet: ' + SHEET_EMP };
    const data       = sheet.getDataRange().getValues();
    const searchCode = String(empCode).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchCode) {
        return {
          code        : data[i][0],
          nameCN      : data[i][2],
          name        : data[i][3],
          department  : data[i][5],
          position    : data[i][10],
          status      : data[i][11],
          affiliation : data[i][15] || '',  // คอลัมน์ P = index 15
        };
      }
    }
    return null;
  } catch(e) {
    return { error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// บันทึกการเบิก + อัพเดตสต็อก
// ══════════════════════════════════════════════════════════════
function saveRequisition(data) {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const tz  = Session.getScriptTimeZone();
    const now = new Date();

    const pSheet = ss.getSheetByName(SHEET_PRODUCT);
    let newStock  = Number(data.stockBefore) || 0;

    if (pSheet) {
      const pData = pSheet.getDataRange().getValues();
      for (let i = 2; i < pData.length; i++) {
        if (String(pData[i][0]).trim() === String(data.productCode).trim()) {
          const cur = Number(pData[i][10]);
          const qty = Number(data.qty);
          if      (data.type === 'เบิกออก')  newStock = cur - qty;
          else if (data.type === 'รับเข้า')   newStock = cur + qty;
          else                                 newStock = cur;
          pSheet.getRange(i + 1, 11).setValue(newStock);
          break;
        }
      }
    }

    const fSheet = ss.getSheetByName(SHEET_FORM);
    if (!fSheet) return { success: false, error: 'ไม่พบ Sheet: ' + SHEET_FORM };

    const cfg           = getAdminConfig();
    const approvalStatus = cfg.requireApproval === 'true' ? 'รอ' : 'อนุมัติ';
    const docId          = generateDocId();
    const dateStr        = data.date || Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    const yearMonth      = Utilities.formatDate(now, tz, 'yyyy-MM');

    fSheet.appendRow([
      docId,              // A
      now,                // B
      dateStr,            // C
      yearMonth,          // D
      '',                 // E
      data.productCode,   // F
      data.productName,   // G
      data.productSize,   // H
      data.qty,           // I
      data.unit,          // J
      '',                 // K
      newStock,           // L
      data.type,          // M
      '',                 // N (ลายเซ็น)
      data.empCode,       // O
      data.empName,       // P
      '',                 // Q
      data.department,    // R
      data.note,          // S
      data.extra || '',   // T
      approvalStatus,     // U ← สถานะอนุมัติ
      data.location,      // V
    ]);

    return { success: true, docId, newStock, approvalStatus };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function generateDocId() {
  const c = 'abcdefghijklmnopqrstuvwxyz0123456789';
  let id = '';
  for (let i = 0; i < 5; i++) id += c[Math.floor(Math.random()*c.length)];
  return id;
}

// ══════════════════════════════════════════════════════════════
// ประวัติ 50 รายการล่าสุด
// ══════════════════════════════════════════════════════════════
function getRecentHistory() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    const tz   = Session.getScriptTimeZone();
    const data = sheet.getRange(3, 1, lastRow - 2, 22).getValues();

    return data.reverse().slice(0, 50).map(r => {
      let dateStr = '';
      const d = r[2];
      if (d instanceof Date && !isNaN(d)) {
        dateStr = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
      } else {
        dateStr = String(d || '');
      }
      return {
        docId      : String(r[0]  || ''),
        date       : dateStr,
        type       : String(r[12] || ''),
        productCode: String(r[5]  || ''),
        productName: String(r[6]  || ''),
        productSize: String(r[7]  || ''),
        qty        : Number(r[8]  || 0),
        unit       : String(r[9]  || ''),
        stock      : String(r[11] || ''),
        empCode    : String(r[14] || '').padStart(10, '0'),
        empName    : String(r[15] || ''),
        department : String(r[17] || ''),
        note       : String(r[18] || ''),
        approval   : String(r[20] || ''),  // col U
        rowNum     : 0,  // ไม่มีใน getRecentHistory เพราะ reverse แล้ว
      };
    });
  } catch(e) {
    Logger.log('getRecentHistory ERROR: ' + e.message);
    return [];
  }
}

// ══════════════════════════════════════════════════════════════
// Dashboard summary
// ══════════════════════════════════════════════════════════════
function getDashboardSummary() {
  try {
    const ss     = SpreadsheetApp.openById(SHEET_ID);
    const pSheet = ss.getSheetByName(SHEET_PRODUCT);
    const fSheet = ss.getSheetByName(SHEET_FORM);
    const tz     = Session.getScriptTimeZone();
    const today  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

    let totalProducts = 0, lowStock = 0, outOfStock = 0, todayReq = 0;

    if (pSheet && pSheet.getLastRow() > 2) {
      const pData = pSheet.getDataRange().getValues().slice(2);
      totalProducts = pData.filter(r => r[0]).length;
      lowStock      = pData.filter(r => r[0] && Number(r[10]) > 0 && Number(r[10]) <= 5).length;
      outOfStock    = pData.filter(r => r[0] && Number(r[10]) <= 0).length;
    }

    if (fSheet && fSheet.getLastRow() > 2) {
      const fData = fSheet.getDataRange().getValues().slice(2);
      todayReq = fData.filter(r => {
        if (!r[0]) return false;
        const d  = r[2];
        const ds = d instanceof Date
          ? Utilities.formatDate(d, tz, 'dd/MM/yyyy')
          : String(d);
        return ds === today;
      }).length;
    }

    return { totalProducts, lowStock, outOfStock, todayReq };
  } catch(e) {
    return { totalProducts:0, lowStock:0, outOfStock:0, todayReq:0 };
  }
}

// ══════════════════════════════════════════════════════════════
// ADMIN — ชีต Admin (สร้างอัตโนมัติถ้ายังไม่มี)
// ══════════════════════════════════════════════════════════════
function _getAdminSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Admin');
  if (!sh) {
    sh = ss.insertSheet('Admin');
    sh.getRange(1,1,1,2).setValues([['KEY','VALUE']]);
    sh.getRange(2,1,10,2).setValues([
      ['adminPin',        '1234'],
      ['lineUrl',         ''],
      ['lineDisplay',     'ติดต่อสโตร์'],
      ['showStock',       'false'],
      ['requireApproval', 'false'],
      ['announcement',    ''],
      ['appName',         'ระบบเบิกอุปกรณ์สำนักงาน'],
      ['companyName',     'บริษัทเคยี่ ก่อสร้าง'],
      ['maxHistoryRows',  '50'],
      ['stockVisibleAll', 'false'],
    ]);
  }
  return sh;
}

function getAdminConfig() {
  const sh   = _getAdminSheet();
  const data = sh.getDataRange().getValues();
  const cfg  = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) cfg[String(data[i][0])] = String(data[i][1]);
  }
  return cfg;
}

function saveAdminConfig(cfg) {
  const sh   = _getAdminSheet();
  const data = sh.getDataRange().getValues();
  for (const key in cfg) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i+1, 2).setValue(cfg[key]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, cfg[key]]);
  }
  return { success: true };
}

function verifyAdminPin(pin) {
  const cfg = getAdminConfig();
  return { ok: String(pin) === String(cfg.adminPin || '1234') };
}

function getPublicConfig() {
  const cfg = getAdminConfig();
  return {
    lineUrl:        cfg.lineUrl         || '',
    lineDisplay:    cfg.lineDisplay     || 'ติดต่อสโตร์',
    showStock:      cfg.showStock       === 'true',
    stockVisibleAll:cfg.stockVisibleAll === 'true',
    requireApproval:cfg.requireApproval === 'true',
    announcement:   cfg.announcement    || '',
    appName:        cfg.appName         || 'ระบบเบิกอุปกรณ์สำนักงาน',
    companyName:    cfg.companyName     || 'บริษัทเคยี่ ก่อสร้าง',
    maxHistoryRows: parseInt(cfg.maxHistoryRows) || 50,
  };
}

// ══════════════════════════════════════════════════════════════
// ประวัติพนักงานรายคน 20 รายการล่าสุด แยกประเภท
// ══════════════════════════════════════════════════════════════
function getEmployeeHistory(empCode) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet || sheet.getLastRow() < 3) return { all:[], out:[], in_:[], loan:[] };

    const tz         = Session.getScriptTimeZone();
    const data       = sheet.getRange(3, 1, sheet.getLastRow()-2, 22).getValues();
    const searchCode = String(empCode).trim().replace(/^0+/,'');
    const result     = [];

    for (let i = data.length - 1; i >= 0; i--) {
      const rowCode = String(data[i][14] || '').trim().replace(/^0+/,'');
      if (rowCode !== searchCode) continue;

      const d = data[i][2];
      const dateStr = d instanceof Date
        ? Utilities.formatDate(d, tz, 'dd/MM/yyyy')
        : String(d || '');

      result.push({
        docId:       String(data[i][0]  || ''),
        date:        dateStr,
        type:        String(data[i][12] || ''),
        productCode: String(data[i][5]  || ''),
        productName: String(data[i][6]  || ''),
        productSize: String(data[i][7]  || ''),
        qty:         Number(data[i][8]  || 0),
        unit:        String(data[i][9]  || ''),
        stock:       String(data[i][11] || ''),
        empCode:     String(data[i][14] || ''),
        empName:     String(data[i][15] || ''),
        department:  String(data[i][17] || ''),
        note:        String(data[i][18] || ''),
        approval:    String(data[i][20] || ''),
        rowNum:      i + 3,
      });
      if (result.length >= 20) break;
    }

    return {
      all:  result,
      out:  result.filter(r => r.type === 'เบิกออก'),
      in_:  result.filter(r => r.type === 'รับเข้า'),
      loan: result.filter(r => r.type === 'ยืมอุปกรณ์'),
    };
  } catch(e) {
    return { all:[], out:[], in_:[], loan:[], error: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// แก้ไข / ลบ / อนุมัติ รายการ
// ══════════════════════════════════════════════════════════════
function updateHistoryRow(rowNum, fields) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet) return { error: 'ไม่พบชีต' };
    const row = sheet.getRange(rowNum, 1, 1, 22).getValues()[0];
    const map = {
      type:12, productCode:5, productName:6, productSize:7,
      qty:8, unit:9, empCode:14, empName:15, department:17, note:18,
    };
    for (const key in fields) {
      if (map[key] !== undefined) row[map[key]] = fields[key];
    }
    sheet.getRange(rowNum, 1, 1, 22).setValues([row]);
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function deleteHistoryRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet) return { error: 'ไม่พบชีต' };
    sheet.deleteRow(rowNum);
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function approveHistoryRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet) return { error: 'ไม่พบชีต' };
    sheet.getRange(rowNum, 21).setValue('อนุมัติ');
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function rejectHistoryRow(rowNum) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FORM);
    if (!sheet) return { error: 'ไม่พบชีต' };
    sheet.getRange(rowNum, 21).setValue('ไม่อนุมัติ');
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

// ══════════════════════════════════════════════════════════════
// สต็อกสินค้าทั้งหมด (หน้าสต็อก Admin)
// ══════════════════════════════════════════════════════════════
function getAllProducts() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRODUCT);
    if (!sheet || sheet.getLastRow() < 3) return [];
    const data = sheet.getDataRange().getValues().slice(2);
    return data
      .filter(r => r[0])
      .map(r => ({
        code    : String(r[0]  || ''),
        category: String(r[2]  || ''),
        name    : String(r[4]  || ''),
        size    : String(r[5]  || ''),
        stock   : Number(r[10] || 0),
        unit    : String(r[11] || ''),
      }));
  } catch(e) { return []; }
}
// 1. เพิ่มฟังก์ชัน OPTIONS เพื่อตอบรับ CORS Preflight
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// 2. ปรับปรุง doGet ให้รองรับ CORS
function doGet(e) {
  // หากคุณใช้หน้า Index.html ต้องเช็คว่ามันเรียกผ่าน GAS หรือเรียกหน้าเว็บตรงๆ
  // สำหรับ API calls ที่ไม่ใช่ HTML, ให้ใช้ ContentService
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเบิกอุปกรณ์สำนักงาน บริษัทเคยี่ ก่อสร้าง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 3. ปรับปรุงการส่งคืนค่าในฟังก์ชันต่างๆ เช่น saveRequisition
// ทุกครั้งที่ return ค่า ให้เพิ่ม .setHeader("Access-Control-Allow-Origin", "*")
function saveRequisition(data) {
  // ... โค้ดเดิมของคุณ ...
  // ส่วนสุดท้ายเปลี่ยนเป็น:
  return ContentService.createTextOutput(JSON.stringify({ success: true, docId, newStock, approvalStatus }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ══════════════════════════════════════════════════════════════
// ทดสอบ — รันใน Apps Script Editor
// ══════════════════════════════════════════════════════════════
function doPost(e) {
  // รับค่าพารามิเตอร์ที่ส่งมาจาก fetch
  var params = JSON.parse(e.postData.contents);
  var action = params.action;
  var data = params.data;
  var result = {};

  // เลือกฟังก์ชันตาม action ที่ส่งมา
  switch (action) {
    case 'saveRequisition':
      result = saveRequisition(data);
      break;
    case 'getProductByCode':
      result = getProductByCode(data.code);
      break;
    case 'getEmployeeByCode':
      result = getEmployeeByCode(data.code);
      break;
    case 'getRecentHistory':
      result = getRecentHistory();
      break;
    case 'getDashboardSummary':
      result = getDashboardSummary();
      break;
    default:
      result = { error: 'ไม่พบคำสั่งที่ระบุ' };
  }

  // ส่งข้อมูลกลับแบบ JSON พร้อม header CORS
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}
function testAll() {
  Logger.log('=== getProductByCode ===');
  Logger.log(JSON.stringify(getProductByCode('8851907335281')));

  Logger.log('=== getEmployeeByCode ===');
  Logger.log(JSON.stringify(getEmployeeByCode('0822920903')));

  Logger.log('=== getRecentHistory ===');
  const h = getRecentHistory();
  Logger.log('จำนวนแถว: ' + h.length);
  if (h.length) Logger.log(JSON.stringify(h[0]));

  Logger.log('=== getDashboardSummary ===');
  Logger.log(JSON.stringify(getDashboardSummary()));

  Logger.log('=== getPublicConfig ===');
  Logger.log(JSON.stringify(getPublicConfig()));

  Logger.log('=== verifyAdminPin ===');
  Logger.log(JSON.stringify(verifyAdminPin('1234')));

  Logger.log('=== getEmployeeHistory ===');
  const eh = getEmployeeHistory('0822920903');
  Logger.log('all:' + eh.all.length + ' out:' + eh.out.length + ' in:' + eh.in_.length + ' loan:' + eh.loan.length);

  Logger.log('=== getAllProducts ===');
  const p = getAllProducts();
  Logger.log('จำนวนสินค้า: ' + p.length);
  if (p.length) Logger.log(JSON.stringify(p[0]));
}
