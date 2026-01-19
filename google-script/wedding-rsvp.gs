function doPost(e) {
  try {
    // Google Sheets ID
    const SHEET_ID = '1ODxRpeg5fbEOsBt0nUswSlZYtsIQwhKL7D9nCl41GFI';
    
    // เปิด Google Sheets
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    
    // ตรวจสอบว่ามี header row หรือไม่ ถ้าไม่มีให้สร้าง
    if (sheet.getLastRow() === 0) {
      const headers = [
        'วันที่/เวลา',
        'ชื่อ-นามสกุล',
        'สถานะการเข้าร่วม',
        'จำนวนผู้เข้าร่วม',
        'ข้อจำกัดด้านอาหาร',
        'คำอวยพร'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // จัดรูปแบบ header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#ec4899');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // รับข้อมูลจากฟอร์ม
    const timestamp = e.parameter.timestamp || new Date().toLocaleString('th-TH');
    const fullName = e.parameter.fullName || '';
    const attendance = e.parameter.attendance || '';
    const guestCount = e.parameter.guestCount || '';
    const dietaryRestrictions = e.parameter.dietaryRestrictions || '';
    const wishes = e.parameter.wishes || '';
    
    // เพิ่มข้อมูลลงใน sheet
    const newRow = [
      timestamp,
      fullName,
      attendance,
      guestCount,
      dietaryRestrictions,
      wishes
    ];
    
    sheet.appendRow(newRow);
    
    // จัดรูปแบบแถวใหม่
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, newRow.length);
    
    // สีพื้นหลังสลับกัน
    if (lastRow % 2 === 0) {
      range.setBackground('#fdf2f8');
    } else {
      range.setBackground('#ffffff');
    }
    
    // จัดการคอลัมน์
    range.setBorder(true, true, true, true, true, true);
    range.setVerticalAlignment('middle');
    
    // ปรับขนาดคอลัมน์อัตโนมัติ
    sheet.autoResizeColumns(1, newRow.length);
    
    // ส่งการตอบกลับ
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'ข้อมูลถูกบันทึกเรียบร้อยแล้ว'
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error:', error);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'เกิดข้อผิดพลาด: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      message: 'RSVP API is working'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}