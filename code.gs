const OUTBOUND_SHEET_NAME = "Outbound_Books";
const CONTROL_SHEET_NAME = "Control_Ledger";
const EVIDENCE_FOLDER_ID = "root";
// นำบรรทัดนี้ไปเพิ่มไว้ด้านบนสุดของไฟล์
const DB_ID = '1nhhXnDwYw1Gaysp50Vew2_-LFsT23oGRkXKIrxQoS2w';

function doGet() {
  ensureDatabaseExists();
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบทะเบียนรับและติดตามหนังสือ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ตรวจสอบและสร้างฐานข้อมูลทั้งหมดของระบบ (เพิ่มโครงสร้างคลังน้ำมัน ยานพาหนะ และระบบใช้รถ)
function ensureDatabaseExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ตาราง Book_Register
  var sheetRegister = ss.getSheetByName('Book_Register');
  var registerHeaders = ['book_id', 'วันที่', 'ทะเบียนรับ', 'กห', 'ที่', 'ลงวันที่', 'จาก', 'เรื่อง', 'status', 'current_dept', 'price', 'book_type'];
  if (!sheetRegister) {
    sheetRegister = ss.insertSheet('Book_Register');
    sheetRegister.appendRow(registerHeaders);
    sheetRegister.getRange(1, 1, 1, registerHeaders.length).setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 2. ตาราง Book_Timeline
  var sheetTimeline = ss.getSheetByName('Book_Timeline');
  var timelineHeaders = ['timeline_id', 'book_id', 'department_from', 'department_to', 'receiver_name', 'signature_data', 'timestamp', 'note'];
  if (!sheetTimeline) {
    sheetTimeline = ss.insertSheet('Book_Timeline');
    sheetTimeline.appendRow(timelineHeaders);
    sheetTimeline.getRange(1, 1, 1, timelineHeaders.length).setFontWeight('bold').setBackground('#065f46').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 3. ตาราง Users
  var sheetUsers = ss.getSheetByName('Users');
  var usersHeaders = ['username', 'password', 'department', 'name_title', 'profile_pic'];
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet('Users');
    sheetUsers.appendRow(usersHeaders);
    sheetUsers.appendRow(['admin', '1234', 'ผกบ.', 'แอดมิน ผกบ.', '']);
    sheetUsers.getRange(1, 1, 1, usersHeaders.length).setFontWeight('bold').setBackground('#9333ea').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 4. ตาราง Settings_Depts
  var sheetDepts = ss.getSheetByName('Settings_Depts');
  if (!sheetDepts) {
    sheetDepts = ss.insertSheet('Settings_Depts');
    sheetDepts.appendRow(['department_name']);
    var defaultDepts = [['ผู้บังคับบัญชา'], ['ผกบ.'], ['ฝพธ.'], ['ฝสภ.'], ['แหล่งรวมรถ'], ['โรงซักรีด'], ['โรงประกอบเลี้ยงคนไข้']];
    sheetDepts.getRange(2, 1, defaultDepts.length, 1).setValues(defaultDepts);
    sheetDepts.getRange(1, 1, 1, 1).setFontWeight('bold').setBackground('#374151').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 5. ตาราง Outbound_Books
  var sheetOutbound = ss.getSheetByName(OUTBOUND_SHEET_NAME);
  var outboundHeaders = ['เลขที่', 'วันที่', 'เรื่อง', 'ไปยัง', 'ประเภท', 'จำนวนรายการ', 'จำนวนราคา', 'ผู้รับผิดชอบ', 'ผู้ตรวจ', 'Timestamp', 'ไฟล์หลักฐาน', 'out_id', 'สป_สาย'];
  if (!sheetOutbound) {
    sheetOutbound = ss.insertSheet(OUTBOUND_SHEET_NAME);
    sheetOutbound.appendRow(outboundHeaders);
    sheetOutbound.getRange(1, 1, 1, outboundHeaders.length).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // 6. ตาราง Control_Ledger
  var sheetControl = ss.getSheetByName(CONTROL_SHEET_NAME);
  var controlHeaders = ['เลขที่', 'วันที่', 'เรื่อง', 'ไปยัง', 'ประเภท', 'จำนวนรายการ', 'จำนวนราคา', 'ผู้รับผิดชอบ', 'ผู้ตรวจ', 'ไฟล์หลักฐาน', 'วันที่ขึ้นบัญชี', 'out_id', 'สป_สาย'];
  if (!sheetControl) {
    sheetControl = ss.insertSheet(CONTROL_SHEET_NAME);
    sheetControl.appendRow(controlHeaders);
    sheetControl.getRange(1, 1, 1, controlHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff').setHorizontalAlignment('center');
  }

  // ตรวจสอบโครงสร้างคลังน้ำมัน ยานพาหนะ และระบบใช้รถประจำวัน
  getOrCreateSheet('Fuel_Depot');
  getOrCreateSheet('Vehicles');
  getOrCreateSheet('Fuel_Log');
  getOrCreateSheet('Vehicle_Usage');
}


// ==========================================
// ฟังก์ชันสำหรับสั่งสร้างฐานข้อมูลครั้งแรก (รันแค่ครั้งเดียว)
// ==========================================
function setupDatabase() {
  try {
    getOrCreateSheet('Fuel_Depot');
    getOrCreateSheet('Vehicles');
    getOrCreateSheet('Fuel_Log');
    getOrCreateSheet('Vehicle_Usage');
    Logger.log("สร้างฐานข้อมูลสำเร็จทั้ง 4 ชีต!");
  } catch (error) {
    Logger.log("เกิดข้อผิดพลาด: " + error);
  }
}

// ฟังก์ชันหลักที่ใช้สร้างชีต
function getOrCreateSheet(sheetName) {
  // ดึงไฟล์ Database ตาม DB_ID ที่ประกาศไว้บนสุดของโปรเจกต์
  var ss = SpreadsheetApp.openById(DB_ID); 
  
  // (ถ้าคุณฝังสคริปต์ไว้ในไฟล์ Google Sheets ให้ใช้บรรทัดล่างนี้แทน)
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    if (sheetName === 'Fuel_Depot') {
      sheet.appendRow(["รหัสใบเบิก", "วันที่", "ปีงบประมาณ", "อัตรา/เครดิต", "ชนิดน้ำมัน", "จำนวน (ลิตร)", "ไฟล์หลักฐาน", "ผู้รับผิดชอบ", "สถานะ"]);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#d97706").setFontColor("#ffffff");
      
    } else if (sheetName === 'Vehicles') {
      sheet.appendRow(["รหัสรถ", "ประเภทรถ", "ยี่ห้อ", "ทะเบียนราชการ", "ทะเบียนพลเรือน", "น้ำมันอัตราพิกัด", "อัตราความสิ้นเปลือง", "เลขไมล์ปัจจุบัน", "ผู้รับผิดชอบ", "เลขไมล์เริ่มต้น"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
      
    } else if (sheetName === 'Fuel_Log') {
      sheet.appendRow(["รหัสการเติม", "วันที่", "รหัสรถ", "ทะเบียนรถ", "ชนิดน้ำมัน", "จำนวน (ลิตร)", "เลขไมล์ตอนเติม", "ผู้บันทึก"]);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
      
    } else if (sheetName === 'Vehicle_Usage') {
      sheet.appendRow(["รหัสการใช้รถ", "รหัสรถ", "ทะเบียนรถ", "พลขับ", "ภารกิจ", "เลขไมล์ก่อนออก", "เลขไมล์กลับ", "เวลาออก", "เวลากลับ", "สถานะ"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    }
    
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

// ----------------------------------------------------
// ระบบลงทะเบียน และยืนยันตัวตน
// ----------------------------------------------------
function getDepartments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings_Depts');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var depts = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) depts.push(data[i][0]); }
  return depts;
}

function registerUser(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0] === d.username) return { success: false, message: "Username ซ้ำ" }; }
  s.appendRow([d.username, d.password, d.department, d.name_title, d.profile_pic]);
  return { success: true, message: "สมัครสำเร็จ" };
}

function verifyLogin(u, p) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == u && data[i][1] == p) return { success: true, username: data[i][0], password: data[i][1], department: data[i][2], name: data[i][3], profile_pic: data[i][4] };
  }
  return { success: false, message: "ชื่อผู้ใช้/รหัสผ่านผิด" };
}

function updateProfile(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === d.username) {
      s.getRange(i + 1, 2).setValue(d.password); s.getRange(i + 1, 4).setValue(d.name_title);
      if (d.profile_pic) s.getRange(i + 1, 5).setValue(d.profile_pic);
      return { success: true, message: "อัปเดตสำเร็จ", name: d.name_title, profile_pic: d.profile_pic || data[i][4] };
    }
  }
}

function getBooks(dept) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Register');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var books = [];
  for (var i = 1; i < data.length; i++) {
    if (dept === 'ผกบ.' || data[i][9] === dept) {
      var dateVal = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : (data[i][1] || "");
      var bookDateVal = (data[i][5] instanceof Date) ? Utilities.formatDate(data[i][5], "GMT+7", "yyyy-MM-dd") : (data[i][5] || "");
      books.push({
        book_id: data[i][0] || "", วันที่: dateVal, ทะเบียนรับ: data[i][2] || "", กห: data[i][3] || "", ที่: data[i][4] || "", ลงวันที่: bookDateVal,
        จาก: data[i][6] || "", เรื่อง: data[i][7] || "", status: data[i][8] || "รอดำเนินการ", current_dept: data[i][9] || "", price: data[i][10] || "", book_type: data[i][11] || "ทั่วไป"
      });
    }
  }
  return books.reverse();
}

function forwardBook(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName('Book_Register');
  var data = regSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.book_id) {
      var targetStatus = d.is_last ? 'เสร็จสิ้น' : (d.is_return ? 'ส่งคืนแก้ไข' : 'อยู่ระหว่างดำเนินการ');
      regSheet.getRange(i + 1, 9).setValue(targetStatus); regSheet.getRange(i + 1, 10).setValue(d.department_to);
      break;
    }
  }
  var prefixNote = d.is_return ? "[ส่งคืนแก้ไข] " : "";
  var timelineNote = prefixNote + "เลขทะเบียนรับหน่วย: " + (d.regNum || "-") + " | วันที่รับ: " + (d.regDate || "-");
  if (d.note) { timelineNote += " | หมายเหตุ: " + d.note; }
  ss.getSheetByName('Book_Timeline').appendRow(['TL-' + new Date().getTime(), d.book_id, d.department_from, d.department_to, d.receiver_name, d.signature_data, new Date(), timelineNote]);
  return true;
}

function toggleReturnStatus(bookId, isReturn, userDept, userName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName('Book_Register');
  var data = regSheet.getDataRange().getValues();
  var newStatus = isReturn ? 'ส่งคืนแก้ไข' : 'อยู่ระหว่างดำเนินการ';
  for (var i = 1; i < data.length; i++) { if (data[i][0] == bookId) { regSheet.getRange(i + 1, 9).setValue(newStatus); break; } }
  ss.getSheetByName('Book_Timeline').appendRow(['TL-' + new Date().getTime(), bookId, userDept, userDept, userName, '', new Date(), isReturn ? '[สลับสวิตช์ตาราง] เปลี่ยนสถานะเป็นส่งคืนกลับไปแก้ไข' : '[สลับสวิตช์ตาราง] ยกเลิกสถานะส่งคืนแก้ไข (กลับเข้าสู่กระบวนการ)']);
  return true;
}

function registerBook(d) { SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Register').appendRow(['BK-'+new Date().getTime(), d.date, d.regNum, d.kh, d.bookNum, d.bookDate, d.fromUnit, d.subject, 'รอดำเนินการ', d.current_dept, d.price, d.bookType]); }
function updateBook(d) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Register'); var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.book_id) {
      s.getRange(i + 1, 2).setValue(d.date); s.getRange(i + 1, 3).setValue(d.regNum); s.getRange(i + 1, 4).setValue(d.kh);
      s.getRange(i + 1, 5).setValue(d.bookNum); s.getRange(i + 1, 6).setValue(d.bookDate); s.getRange(i + 1, 7).setValue(d.fromUnit);
      s.getRange(i + 1, 8).setValue(d.subject); s.getRange(i + 1, 11).setValue(d.price); s.getRange(i + 1, 12).setValue(d.bookType);
      break;
    }
  }
  return true;
}

function deleteBook(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var s = ss.getSheetByName('Book_Register'); var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0] == id) { s.deleteRow(i + 1); break; } }
  var t = ss.getSheetByName('Book_Timeline'); var td = t.getDataRange().getValues();
  for (var j = td.length - 1; j >= 1; j--) { if (td[j][1] == id) t.deleteRow(j + 1); }
  return true;
}

function getTimeline(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Timeline'); if (!sheet) return [];
  var data = sheet.getDataRange().getValues(); var timelines = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == id) {
      var ts = data[i][6] ? Utilities.formatDate(new Date(data[i][6]), "GMT+7", "dd/MM/yyyy HH:mm") : "";
      timelines.push({ timeline_id: data[i][0], book_id: data[i][1], department_from: data[i][2], department_to: data[i][3], receiver_name: data[i][4], signature_data: data[i][5], timestamp: ts, note: data[i][7] });
    }
  }
  return timelines;
}

function getUserWorkload() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Book_Timeline');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var counts = {};
  for (var i = 1; i < data.length; i++) {
    var name = data[i][4];
    var dept = data[i][2];
    if (name && dept) {
      var key = name + " [" + dept + "]";
      counts[key] = (counts[key] || 0) + 1;
    }
  }
  var list = [];
  for (var key in counts) {
    list.push({ label: key, value: counts[key] });
  }
  list.sort(function(a, b) { return b.value - a.value; });
  return list;
}

function getGlobalTimeline() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timelineSheet = ss.getSheetByName("Book_Timeline") || ss.getSheets()[1];
    if (!timelineSheet) return [];
    var timelineData = timelineSheet.getDataRange().getValues();
    if (timelineData.length <= 1) return [];

    var result = [];
    var startRow = Math.max(1, timelineData.length - 15);
    for (var j = timelineData.length - 1; j >= startRow; j--) {
      var row = timelineData[j];
      result.push({
        department_from: row[2] || "ระบบ",
        department_to: row[3] || "-",
        receiver_name: row[4] || "-",
        timestamp: row[6] ? Utilities.formatDate(new Date(row[6]), "GMT+7", "dd/MM/yyyy HH:mm") : "",
        note: row[7] || "",
        subject: "อัปเดตสถานะเดินหนังสือ"
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

function getNextRegisterNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Book_Register'); 
  return sheet.getLastRow(); 
}

// ----------------------------------------------------
// ระบบสารบรรณขึ้นบัญชีคุม สป. ทั่วไป
// ----------------------------------------------------
function getOutboundBooks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTBOUND_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
    result.push({
      no: data[i][0], date: dStr, subject: data[i][2], to: data[i][3], type: data[i][4],
      itemCount: data[i][5], price: data[i][6], responsible: data[i][7], approver: data[i][8],
      timestamp: data[i][9], file_url: data[i][10], out_id: data[i][11], sapo: data[i][12]
    });
  }
  return result;
}

function saveOutboundBook(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTBOUND_SHEET_NAME);
  if (d.out_id) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][11] == d.out_id) {
        sheet.getRange(i+1, 2, 1, 8).setValues([[d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver]]);
        sheet.getRange(i+1, 13).setValue(d.sapo);
        return { success: true };
      }
    }
  } else {
    var nextNo = sheet.getLastRow();
    var newId = 'OUT-' + new Date().getTime();
    sheet.appendRow([nextNo, d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver, new Date(), '', newId, d.sapo]);
    return { success: true };
  }
}

function deleteOutboundBook(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OUTBOUND_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function getControlLedger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
    var tsStr = (data[i][10] instanceof Date) ? Utilities.formatDate(data[i][10], "GMT+7", "dd/MM/yyyy HH:mm") : data[i][10];
    result.push({
      no: data[i][0], date: dStr, subject: data[i][2], to: data[i][3], type: data[i][4],
      itemCount: data[i][5], price: data[i][6], responsible: data[i][7], approver: data[i][8],
      file_url: data[i][9], timestamp: tsStr, out_id: data[i][11], sapo: data[i][12]
    });
  }
  return result.reverse();
}

function updateControlLedgerEntry(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] == d.out_id) {
      sheet.getRange(i+1, 1, 1, 9).setValues([[d.no, d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver]]);
      sheet.getRange(i+1, 13).setValue(d.sapo);
      return { success: true };
    }
  }
  // กรณีเพิ่มใหม่โดยตรง
  var newId = d.out_id || 'CTL-' + new Date().getTime();
  sheet.appendRow([d.no, d.date, d.subject, d.to, d.type, d.itemCount, d.price, d.responsible, d.approver, '', new Date(), newId, d.sapo]);
  return { success: true };
}

function deleteControlLedgerEntry(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][11] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function registerToControlLedger(outId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetOut = ss.getSheetByName(OUTBOUND_SHEET_NAME);
  var sheetCtrl = ss.getSheetByName(CONTROL_SHEET_NAME);
  var dataOut = sheetOut.getDataRange().getValues();
  for (var i = 1; i < dataOut.length; i++) {
    if (dataOut[i][11] == outId) {
      var nextNo = sheetCtrl.getLastRow();
      sheetCtrl.appendRow([nextNo, dataOut[i][1], dataOut[i][2], dataOut[i][3], dataOut[i][4], dataOut[i][5], dataOut[i][6], dataOut[i][7], dataOut[i][8], dataOut[i][10], new Date(), outId, dataOut[i][12]]);
      return { success: true };
    }
  }
  return { success: false, message: "ไม่พบข้อมูลต้นทาง" };
}

// ----------------------------------------------------
// ระบบคลังน้ำมันและยานพาหนะ (Fuel System)
// ----------------------------------------------------
function getFuelDepots() {
  try {
    var sheet = getOrCreateSheet('Fuel_Depot');
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
      result.push({
        id: String(data[i][0]), date: dStr, year: String(data[i][2]), credit: String(data[i][3]),
        type: String(data[i][4]), liter: String(data[i][5]), file: String(data[i][6]),
        responsible: String(data[i][7]), status: String(data[i][8])
      });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function saveFuelDepot(d) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.id) {
      sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.year, d.credit, d.type, d.liter, d.file || '', d.responsible]]);
      sheet.getRange(i+1, 9).setValue(d.status);
      return { success: true };
    }
  }
  sheet.appendRow([d.id, d.date, d.year, d.credit, d.type, d.liter, d.file || '', d.responsible, d.status]);
  return { success: true };
}

function deleteFuelDepot(id) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

function updateFuelDepotStatus(id, newStatus) {
  var sheet = getOrCreateSheet('Fuel_Depot');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.getRange(i + 1, 9).setValue(newStatus); return { success: true }; }
  }
  return { success: false };
}

// ระบบจัดการยานพาหนะพร้อมอัตราสิ้นเปลือง และคำนวนน้ำมันคงถังสะสมแบบเรียลไทม์
function getVehicles() {
  try {
    var sheetV = getOrCreateSheet('Vehicles');
    var sheetL = getOrCreateSheet('Fuel_Log');
    var vData = sheetV.getDataRange().getValues();
    var lData = sheetL.getDataRange().getValues();
    
    var result = [];
    for (var i = 1; i < vData.length; i++) {
      if (!vData[i][0]) continue;
      var vId = String(vData[i][0]);
      
      // 1. คำนวณยอดเติมน้ำมันสะสม
      var totalFill = 0;
      for (var j = 1; j < lData.length; j++) {
        if (String(lData[j][2]) === vId) {
          totalFill += Number(lData[j][5]) || 0;
        }
      }

      // 2. จัดการเลขไมล์เริ่มต้น (ดึงจากคอลัมน์ J)
      var currMile = Number(vData[i][7]) || 0; // คอลัมน์ H (ไมล์ปัจจุบัน)
      var startMileRaw = vData[i][9];          // คอลัมน์ J (ไมล์เริ่มต้น)
      var startMile = Number(startMileRaw);
      
      // ถ้าไม่มีเลขไมล์เริ่มต้น โค้ดจะดึงไมล์ปัจจุบันไปใช้ และเขียนลงฐานข้อมูลให้อัตโนมัติ
      if (startMileRaw === "" || startMileRaw === undefined) {
        startMile = currMile;
        sheetV.getRange(i + 1, 10).setValue(startMile);
      }

      var distance = Math.max(0, currMile - startMile);
      
      // 3. ประเมินน้ำมันคงเหลือในถัง
      var consumeRate = Number(vData[i][6]) || 1;
      var estimatedConsume = distance / consumeRate;
      var fuelInTank = Math.max(0, Math.round(totalFill - estimatedConsume));

      // ส่งข้อมูลออกไปหน้าเว็บ
      result.push({
        id: vId, type: String(vData[i][1]), brand: String(vData[i][2]),
        govReg: String(vData[i][3]), civilReg: String(vData[i][4]),
        quota: String(vData[i][5]), consume: String(vData[i][6]),
        mileage: String(vData[i][7]), driver: String(vData[i][8]),
        startMileage: String(startMile), distance: distance, fuelInTank: fuelInTank
      });
    }
    return result;
  } catch (e) { return []; }
}

function saveVehicle(d) {
  try {
    var sheet = getOrCreateSheet('Vehicles');
    var data = sheet.getDataRange().getValues();
    
    // ตั้งค่าเลขไมล์เริ่มต้น: ถ้าหน้าเว็บส่ง d.startMileage มาให้ใช้ค่านั้น ถ้าไม่ส่งมาให้ใช้ d.mileage แทน
    var startMile = d.startMileage ? d.startMileage : d.mileage;
    
    // 1. กรณี: แก้ไขข้อมูลรถเดิม
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == d.id) {
        // อัปเดตรวดเดียวตั้งแต่คอลัมน์ที่ 2 (ประเภทรถ) จนถึงคอลัมน์ที่ 10 (เลขไมล์เริ่มต้น)
        sheet.getRange(i+1, 2, 1, 9).setValues([[
          d.type, d.brand, d.govReg, d.civilReg, d.quota, d.consume, d.mileage, d.driver, startMile
        ]]);
        return { success: true };
      }
    }
    
    // 2. กรณี: เพิ่มรถใหม่
    var id = "V" + new Date().getTime();
    sheet.appendRow([
      id, d.type, d.brand, d.govReg, d.civilReg, d.quota, d.consume, d.mileage, d.driver, startMile
    ]);
    
    return { success: true };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteVehicle(id) {
  var sheet = getOrCreateSheet('Vehicles');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log) ปรับปรุงใหม่
// ==========================================

/**
 * ดึงข้อมูลประวัติการเติมน้ำมันทั้งหมด (เชื่อมต่อผ่าน DB_ID)
 */


/**
 * บันทึก/แก้ไขประวัติการเติมน้ำมัน
 */
function saveFuelLog(d) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var isEdit = false;
    
    // ตรวจสอบว่าเป็นการแก้ไขข้อมูลเดิมหรือไม่
    if (d.logId) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == d.logId) {
          sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]]);
          isEdit = true; 
          break;
        }
      }
    }
    
    // หากไม่ใช่การแก้ไข ให้เพิ่มแถวใหม่
    if (!isEdit) {
      var logId = "L" + new Date().getTime(); // สร้าง ID อัตโนมัติ
      sheet.appendRow([logId, d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]);
    }
    
    // อัปเดตเลขไมล์กลับไปที่ฐานข้อมูล Vehicles
    var sheetV = getOrCreateSheet('Vehicles');
    var vData = sheetV.getDataRange().getValues();
    for (var k = 1; k < vData.length; k++) {
      if (vData[k][0] == d.vehicleId) {
        var currentMileage = Number(vData[k][7] || 0);
        // เช็กไมล์ตอนเติม ต้องมากกว่าเลขไมล์ปัจจุบันในระบบถึงจะอัปเดต
        if (Number(d.mileage) > currentMileage) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.mileage));
        }
        break;
      }
    }
    
    return { success: true, message: 'บันทึกประวัติการเติมน้ำมันเรียบร้อย' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * ลบประวัติการเติมน้ำมัน
 */
function deleteFuelLog(id) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) { 
        sheet.deleteRow(i + 1); 
        return { success: true, message: 'ลบประวัติการเติมน้ำมันสำเร็จ' }; 
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch (error) {
    return { success: false, message: 'ลบข้อมูลไม่สำเร็จ: ' + error.toString() };
  }
}





function getFuelDashboardOverview() {
  try {
    var stats = { totalVehicles: 0, totalFuelAdded: 0, totalFuelDisbursed: 0, pendingRequests: 0, totalFuelInDepot: 0 };
    var sheetVehicles = getOrCreateSheet('Vehicles');
    stats.totalVehicles = Math.max(0, sheetVehicles.getLastRow() - 1);

    var sheetDepot = getOrCreateSheet('Fuel_Depot');
    var depotData = sheetDepot.getDataRange().getValues();
    for (var i = 1; i < depotData.length; i++) {
      var status = String(depotData[i][8]);
      if (status === 'รออนุมัติจ่าย' || status === 'รอดำเนินการ') {
        stats.pendingRequests++;
      }
      if (status === 'รับน้ำมันแล้ว' || status === 'เสร็จสิ้น') {
        stats.totalFuelAdded += Number(depotData[i][5]) || 0;
      }
    }

    var sheetLog = getOrCreateSheet('Fuel_Log');
    var logData = sheetLog.getDataRange().getValues();
    for (var j = 1; j < logData.length; j++) {
      stats.totalFuelDisbursed += Number(logData[j][5]) || 0;
    }

    stats.totalFuelInDepot = Math.max(0, stats.totalFuelAdded - stats.totalFuelDisbursed);
    return stats;
  } catch (e) {
    return { totalVehicles: 0, totalFuelAdded: 0, totalFuelDisbursed: 0, pendingRequests: 0, totalFuelInDepot: 0 };
  }
}

// ----------------------------------------------------
// ระบบขอใช้รถยนต์ประจำวัน (Vehicle Daily Usage)
// ----------------------------------------------------
function getVehicleUsages() {
  try {
    var sheet = getOrCreateSheet('Vehicle_Usage');
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var depTime = data[i][7] ? Utilities.formatDate(new Date(data[i][7]), "GMT+7", "yyyy-MM-dd'T'HH:mm") : "";
      var retTime = data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+7", "yyyy-MM-dd'T'HH:mm") : "";
      result.push({
        usageId: String(data[i][0]), vehicleId: String(data[i][1]), vehicleReg: String(data[i][2]),
        driver: String(data[i][3]), purpose: String(data[i][4]), startMileage: String(data[i][5]),
        endMileage: String(data[i][6]), departureTime: depTime, returnTime: retTime, status: String(data[i][9])
      });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function saveVehicleUsage(d) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var usageId = "USG" + new Date().getTime();
  sheet.appendRow([usageId, d.vehicleId, d.vehicleReg, d.driver, d.purpose, d.startMileage, '', d.departureTime, '', 'อยู่ระหว่างภารกิจ']);
  
  // อัปเดตไมล์เริ่มต้นให้รถทันที
  var sheetV = getOrCreateSheet('Vehicles');
  var vData = sheetV.getDataRange().getValues();
  for (var i = 1; i < vData.length; i++) {
    if (vData[i][0] == d.vehicleId) {
      sheetV.getRange(i + 1, 8).setValue(Number(d.startMileage)); break;
    }
  }
  return { success: true };
}

function returnVehicleUsage(d) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == d.usageId) {
      sheet.getRange(i + 1, 7).setValue(Number(d.endMileage));
      sheet.getRange(i + 1, 9).setValue(d.returnTime);
      sheet.getRange(i + 1, 10).setValue('เสร็จสิ้นภารกิจ');
      
      // ส่งไมล์สะสมล่าสุดกลับไปที่ Vehicles
      var sheetV = getOrCreateSheet('Vehicles');
      var vData = sheetV.getDataRange().getValues();
      for (var k = 1; k < vData.length; k++) {
        if (vData[k][0] == d.vehicleId) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.endMileage)); break;
        }
      }
      return { success: true };
    }
  }
  return { success: false };
}

function deleteVehicleUsage(id) {
  var sheet = getOrCreateSheet('Vehicle_Usage');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false };
}

// ฟังก์ชันสำหรับแก้ไขข้อมูล
function updateRecord(sheetName, idColumnIndex, recordId, newData) {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    
    // วนลูปหาแถวที่มี ID ตรงกัน (เริ่มที่ 1 เพื่อข้ามหัวตาราง)
    for(let i = 1; i < data.length; i++) {
       if(data[i][idColumnIndex] == recordId) {
          // i+1 เพราะ Array เริ่มที่ 0 แต่ Row เริ่มที่ 1
          sheet.getRange(i+1, 1, 1, newData.length).setValues([newData]);
          return {success: true, message: "อัปเดตข้อมูลสำเร็จ"};
       }
    }
    return {success: false, message: "ไม่พบรหัสข้อมูลที่ต้องการแก้ไข"};
  } catch(e) {
    return {success: false, message: "Error: " + e.message};
  }
}

// ฟังก์ชันสำหรับลบข้อมูล
function deleteRecord(sheetName, idColumnIndex, recordId) {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    
    for(let i = 1; i < data.length; i++) {
       if(data[i][idColumnIndex] == recordId) {
          sheet.deleteRow(i+1);
          return {success: true, message: "ลบข้อมูลสำเร็จ"};
       }
    }
    return {success: false, message: "ไม่พบรหัสข้อมูลที่ต้องการลบ"};
  } catch(e) {
    return {success: false, message: "Error: " + e.message};
  }
}

function getDashboardSummary() {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    
    // ดึงข้อมูลจากชีตต่างๆ
    let vehicleData = ss.getSheetByName('Vehicles').getDataRange().getValues();
    let usageData = ss.getSheetByName('Vehicle_Usage').getDataRange().getValues();
    let fuelLogData = ss.getSheetByName('Fuel_Log').getDataRange().getValues();
    
    let dashboardStats = [];
    
    // ข้ามแถวที่ 0 (หัวตาราง) เริ่มดึงข้อมูลรถทีละคัน
    for (let i = 1; i < vehicleData.length; i++) {
      let vCode = vehicleData[i][0]; // รหัสรถ
      let vReg = vehicleData[i][3];  // ทะเบียนราชการ
      
      let totalDistance = 0;
      let totalFuelAdded = 0;
      
      // 1. คำนวณระยะทางรวมจากชีต Vehicle_Usage
      for (let j = 1; j < usageData.length; j++) {
        if (usageData[j][1] === vCode && usageData[j][6] !== "" && usageData[j][5] !== "") {
          // (เลขไมล์กลับ - เลขไมล์ก่อนออก)
          let distance = Number(usageData[j][6]) - Number(usageData[j][5]);
          if (distance > 0) totalDistance += distance;
        }
      }
      
      // 2. คำนวณน้ำมันที่เติมรวมจากชีต Fuel_Log
      for (let k = 1; k < fuelLogData.length; k++) {
        if (fuelLogData[k][2] === vCode && fuelLogData[k][5] !== "") {
          totalFuelAdded += Number(fuelLogData[k][5]);
        }
      }
      
      // 3. คำนวณอัตราสิ้นเปลือง (Km/L)
      let consumptionRate = 0;
      if (totalFuelAdded > 0) {
        consumptionRate = (totalDistance / totalFuelAdded).toFixed(2);
      }
      
      // เก็บข้อมูลเข้า Array เพื่อส่งไปแสดงที่หน้าเว็บ
      dashboardStats.push({
        vehicleCode: vCode,
        registration: vReg,
        distanceStr: totalDistance + " กม.",
        fuelStr: totalFuelAdded + " ลิตร",
        kmPerLiter: consumptionRate + " กม./ลิตร"
      });
    }
    
    return { success: true, data: dashboardStats };
    
  } catch(e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}

// ฟังก์ชันดึงรายชื่อรถจากชีต Vehicles
function getVehicleList() {
  try {
    let ss = SpreadsheetApp.openById(DB_ID);
    let sheet = ss.getSheetByName('Vehicles');
    let data = sheet.getDataRange().getDisplayValues(); // ใช้ getDisplayValues เพื่อให้ได้ข้อความตรงตามที่เห็นในชีต
    
    let vehicleList = [];
    
    // เริ่มวนลูปที่ i = 1 เพื่อข้ามหัวตารางแถวแรก
    for (let i = 1; i < data.length; i++) {
      let vCode = data[i][0]; // คอลัมน์ A : รหัสรถ
      let vReg = data[i][3];  // คอลัมน์ D : ทะเบียนราชการ
      
      // ถ้ามีรหัสรถ ให้เก็บเข้า Array
      if (vCode !== "") {
        vehicleList.push({
          code: vCode,
          label: vCode + " (" + vReg + ")" // รูปแบบที่จะแสดงใน Dropdown
        });
      }
    }
    
    return vehicleList;
    
  } catch(e) {
    return []; // ถ้ามี Error ให้ส่ง Array ว่างกลับไป
  }
}

// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log)
// ==========================================

/**
 * ดึงข้อมูลประวัติการเติมน้ำมันทั้งหมด
 */
// ==========================================
// ส่วนจัดการข้อมูล: ประวัติการเติมน้ำมัน (Fuel_Log)
// ==========================================

function getFuelLogs() {
  try {
    // บังคับดึงข้อมูลจาก DB_ID
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dStr = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], "GMT+7", "yyyy-MM-dd") : data[i][1];
      result.push({
        logId: String(data[i][0]),
        date: dStr,
        vehicleId: String(data[i][2]),
        vehicleReg: String(data[i][3]),
        type: String(data[i][4]),
        liter: String(data[i][5]),
        mileage: String(data[i][6]),
        recorder: String(data[i][7])
      });
    }
    return result.reverse();
  } catch (e) {
    return [];
  }
}

function saveFuelLog(d) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var isEdit = false;
    
    if (d.logId) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] == d.logId) {
          sheet.getRange(i+1, 2, 1, 7).setValues([[d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]]);
          isEdit = true; 
          break;
        }
      }
    }
    
    if (!isEdit) {
      var logId = "L" + new Date().getTime();
      sheet.appendRow([logId, d.date, d.vehicleId, d.vehicleReg, d.type, d.liter, d.mileage, d.recorder]);
    }
    
    // อัปเดตเลขไมล์กลับไปที่ฐานข้อมูล Vehicles
    var sheetV = getOrCreateSheet('Vehicles');
    var vData = sheetV.getDataRange().getValues();
    for (var k = 1; k < vData.length; k++) {
      if (vData[k][0] == d.vehicleId) {
        var currentMileage = Number(vData[k][7] || 0);
        if (Number(d.mileage) > currentMileage) {
          sheetV.getRange(k + 1, 8).setValue(Number(d.mileage));
        }
        break;
      }
    }
    return { success: true, message: 'บันทึกประวัติการเติมน้ำมันเรียบร้อย' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteFuelLog(id) {
  try {
    var sheet = getOrCreateSheet('Fuel_Log');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) { 
        sheet.deleteRow(i + 1); 
        return { success: true, message: 'ลบข้อมูลสำเร็จ' }; 
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * บันทึกประวัติการเติมน้ำมัน (รองรับทั้งการสร้างใหม่และการแก้ไข)
 */
function saveFuelLog(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Fuel_Log');
    
    if (!sheet) {
      sheet = ss.insertSheet('Fuel_Log');
      sheet.appendRow(['logId', 'date', 'vehicleId', 'vehicleReg', 'type', 'liter', 'mileage', 'recorder', 'timestamp']);
    }

    if (data.logId) {
      // กรณี: แก้ไขข้อมูลเดิม
      const sheetData = sheet.getDataRange().getValues();
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][0] == data.logId) {
          sheet.getRange(i + 1, 2).setValue(data.date);
          sheet.getRange(i + 1, 3).setValue(data.vehicleId);
          sheet.getRange(i + 1, 4).setValue(data.vehicleReg);
          sheet.getRange(i + 1, 5).setValue(data.type);
          sheet.getRange(i + 1, 6).setValue(data.liter);
          sheet.getRange(i + 1, 7).setValue(data.mileage);
          sheet.getRange(i + 1, 8).setValue(data.recorder);
          sheet.getRange(i + 1, 9).setValue(new Date()); // อัปเดตเวลาล่าสุด
          break;
        }
      }
    } else {
      // กรณี: บันทึกข้อมูลใหม่
      const newId = 'FL-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
      sheet.appendRow([
        newId,
        data.date,
        data.vehicleId,
        data.vehicleReg,
        data.type,
        data.liter,
        data.mileage,
        data.recorder,
        new Date()
      ]);
    }

    // ฟังก์ชันเสริม: อัปเดตเลขไมล์ปัจจุบันไปที่ทำเนียบรถยนต์ (Vehicles) ให้อัตโนมัติด้วย
    updateVehicleMileageFromFuelLog(data.vehicleId, data.mileage);

    return { success: true, message: 'บันทึกประวัติการเติมน้ำมันเรียบร้อย' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * ลบประวัติการเติมน้ำมัน
 */
function deleteFuelLog(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Fuel_Log');
    if (!sheet) return { success: false, message: 'ไม่พบฐานข้อมูล Fuel_Log' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบประวัติการเติมน้ำมันสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
  } catch (error) {
    return { success: false, message: 'ลบข้อมูลไม่สำเร็จ: ' + error.toString() };
  }
}

/**
 * [ฟังก์ชันช่วยเหลือ] - นำเลขไมล์ตอนเติมน้ำมันไปอัปเดตเป็นเลขไมล์ปัจจุบันในชีต Vehicles
 */
function updateVehicleMileageFromFuelLog(vehicleId, newMileage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    if (!vSheet) return;

    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (vData[i][0] == vehicleId) {
        // อ้างอิงว่าในชีต Vehicles เลขไมล์ล่าสุดเก็บไว้ที่คอลัมน์ไหน
        // *หมายเหตุ: สมมติว่าเลขไมล์เก็บอยู่คอลัมน์ที่ 8 (Index 7) หากชีตของคุณอยู่คอลัมน์อื่น ให้เปลี่ยนเลขบรรทัดด้านล่าง
        const currentMileage = Number(vData[i][7] || 0);
        
        // อัปเดตเฉพาะกรณีที่เลขไมล์ใหม่ มากกว่าเลขไมล์เดิมในระบบเท่านั้น
        if (Number(newMileage) > currentMileage) {
          vSheet.getRange(i + 1, 8).setValue(newMileage);
        }
        break;
      }
    }
  } catch(e) {
    console.error("Error updating mileage to Vehicles sheet: " + e);
  }
}

// ==========================================
// ฟังก์ชันดึงรายชื่อผู้ตรวจ (ดึงจากชีต Users)
// ==========================================
function getApprovers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const approvers = [];
    
    // เริ่มอ่านข้อมูลบรรทัดที่ 2 เป็นต้นไป (ข้ามหัวตาราง)
    for (let i = 1; i < data.length; i++) {
      // ตรวจสอบว่ามี Username และชื่อ-สกุล
      if (data[i][0] && data[i][3]) {
        approvers.push({ 
          username: data[i][0], 
          name: data[i][3], 
          department: data[i][2] 
        });
      }
    }
    return approvers;
  } catch (error) {
    return [];
  }
}

function getUsers() {
  try {
    // ดึงข้อมูลจากชีตที่ชื่อ 'Users' (ถ้าของคุณชื่ออื่น เปลี่ยนตรงนี้ได้ครับ)
    var sheet = getOrCreateSheet('Users'); 
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var result = [];
    
    // ข้ามแถวที่ 1 (หัวตาราง) เริ่มดึงข้อมูลจากแถวที่ 2
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; 
      
      // ส่งข้อมูลโดยจับคู่ชื่อตัวแปรให้ตรงกับที่หน้าเว็บเรียกร้อง
      result.push({
        username: String(data[i][0]),    // ดึงคอลัมน์ A (ลำดับที่ 0) มาเป็น ID
        name: String(data[i][1] || '-'), // ดึงคอลัมน์ B (ลำดับที่ 1) มาเป็น ชื่อ
        department: String(data[i][2] || '-') // ดึงคอลัมน์ C (ลำดับที่ 2) มาเป็น แผนก
        
        // หมายเหตุ: ถ้าในฐานข้อมูลมีคอลัมน์รูปโปรไฟล์ (เช่น คอลัมน์ D) สามารถเพิ่มบรรทัดนี้ได้:
        // profile_pic: String(data[i][3]) 
      });
    }
    
    return result; 
  } catch (e) {
    return [];
  }
}
