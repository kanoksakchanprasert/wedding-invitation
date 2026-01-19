function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // แก้ไขให้เรียกใช้ Spreadsheet ตาม ID ที่ระบุ
    var ss = SpreadsheetApp.openById('1pB8v0MW6LgDLmQSLEgDBIaARCWJk5a59G7WN0v45HDY');
    var action = e.parameter.action;

    // 1. เริ่มรอบใหม่ (New Round) -> สร้าง Sheet ใหม่ต่อท้าย
    if (action == 'new_round') {
      var nextRoundNum = ss.getSheets().length + 1;
      var newSheetName = "Round " + nextRoundNum;
      
      // สร้าง Sheet ใหม่ไว้ตำแหน่งสุดท้าย
      var newSheet = ss.insertSheet(newSheetName, ss.getNumSheets());
      newSheet.appendRow(["Word", "Timestamp"]); // หัวตาราง
      
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success', 
        round: newSheetName
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. บันทึกคำศัพท์ (Add Word) -> ลง Sheet แผ่นสุดท้าย (รอบปัจจุบัน)
    if (action == 'add') {
      var sheets = ss.getSheets();
      var currentSheet = sheets[sheets.length - 1]; // เลือก Sheet แผ่นสุดท้ายเสมอ
      
      currentSheet.appendRow([e.parameter.word, new Date()]);
      
      return ContentService.createTextOutput(JSON.stringify({status: 'success'})).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}