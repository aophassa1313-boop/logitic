const SPREADSHEET_ID = '1chna_F2hten0RzmUOQyJpT_RVRyJpinUagRlw_SJLbI';

// ฟังก์ชันหลักสำหรับเปิด Web App
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบบริหารจัดการ Logistics')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันสำหรับ Login ตรวจสอบจากชีต "สมาชิก"
function authenticateUser(memberId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('สมาชิก');
    if (!sheet) return { success: false, message: 'ไม่พบฐานข้อมูลสมาชิก' };
    
    const data = sheet.getDataRange().getDisplayValues();
    
    // วนลูปหา รหัสสมาชิก (คอลัมน์ A / index 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === memberId) {
        return { 
          success: true, 
          user: {
            id: data[i][0],
            name: data[i][1], // ยศ-ชื่อ-สกุล (คอลัมน์ B / index 1)
            role: data[i][3]  // ตำแหน่ง
          }
        };
      }
    }
    return { success: false, message: 'รหัสสมาชิกไม่ถูกต้อง' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ฟังก์ชันสำหรับดึงข้อมูลจากชีตที่ระบุ
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: true, message: 'ไม่พบชีต: ' + sheetName };
    
    // ใช้ getDisplayValues เพื่อให้ได้ข้อมูลตามรูปแบบที่แสดงบนจอ (เช่น วันที่)
    const data = sheet.getDataRange().getDisplayValues();
    return { error: false, data: data };
  } catch (error) {
    return { error: true, message: error.toString() };
  }
}

// ฟังก์ชันสำหรับอัปเดตข้อมูลทีละเซลล์ (Inline Editing)
function updateCell(sheetName, rowIndex, colIndex, newValue) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'ไม่พบชีต' };
    
    // rowIndex และ colIndex ส่งมาจาก JS จะเริ่มที่ 0
    // แต่ App Script Range เริ่มที่ 1 ดังนั้นต้อง +1
    sheet.getRange(rowIndex + 1, colIndex + 1).setValue(newValue);
    
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ดึงข้อมูลสรุปสำหรับ Dashboard
function getDashboardSummary() {
  // เพิ่ม 'ยานพาหนะ' เข้าไปในรายการชีตที่ต้องติดตามข้อมูล
  const sheetsToTrack = [
    'ทะเบียนรับหนังสือ', 'เสนอความต้องการ', 'ส่งซ่อม จำหน่าย', 
    'จำหน่าย', 'ขึ้นบัญชีคุม', 'คลัง สป.3', 'เบิก สป.3', 
    'จำหน่าย สป.3', 'ทะเบียนหนังสือออก', 'ยานพาหนะ'
  ];
  
  let summaryData = {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  sheetsToTrack.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getDisplayValues();
      summaryData[sheetName] = data; // ส่งข้อมูลดิบไปกรองและนับที่ฝั่ง Frontend
    }
  });
  
  return summaryData;
}
