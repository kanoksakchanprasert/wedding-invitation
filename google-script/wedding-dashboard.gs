/**
 * AppsScriptCode.js (Secure Version)
 * วิธีการติดตั้ง:
 * 1. นำโค้ดนี้ไปทับของเดิมทั้งหมด
 * 2. กด Deploy > New deployment (สำคัญมาก! ต้อง New deployment ทุกครั้งที่แก้โค้ด)
 * 3. Copy URL ใหม่ (หรือใช้ URL เดิมถ้าเลือกเป็น version ใหม่ทับอันเดิม) ไปใส่ในเว็บ
 */

const SHEET_ID = '1miv2SqSwV8arLQ5u0aEjHb3umGiPtBVkjbl7Rz6xCrE'; 
const SHEET_NAME = 'ซอง';
const APP_PASSWORD = 'Kob&Ying17012026'; // ตั้งรหัสผ่านที่นี่ (Backend Side)

// Column Index (A=0, B=1, ...)
const COL_NAME = 0;     
const COL_SIDE = 1;     
const COL_RELATION = 2; 
const COL_CHANNEL = 3;  
const COL_AMOUNT = 4;   
const COL_GIFT = 5;     
const COL_NOTES = 6;    

function doGet(e) {
  // เพื่อความปลอดภัยและรองรับ payload ขนาดใหญ่ เราจะเน้นใช้ doPost
  // แต่ถ้าเรียกผ่าน Browser ตรงๆ ให้แสดงข้อความบอก
  return ContentService.createTextOutput("Please use POST method for secure data fetching.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!e.postData || !e.postData.contents) {
       return createJSONOutput({ status: 'error', message: 'No data' });
    }

    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const clientPassword = body.password || ""; // รับรหัสจากหน้าเว็บ

    // ตรวจสอบความถูกต้องของรหัสผ่าน
    const isAuthorized = (clientPassword === APP_PASSWORD);

    if (action === 'get_data') {
      // --- READ DATA ---
      const dataRange = sheet.getDataRange().getValues();
      const rows = dataRange.slice(1); // Skip Header
      
      const result = rows.map((r, index) => {
        // ดึงข้อมูลพื้นฐานที่ไม่อันตราย (สำหรับทำ Stats)
        const baseData = {
          row: index + 2,
          side: r[COL_SIDE],
          relation: r[COL_RELATION],
          channel: r[COL_CHANNEL],
          amount: r[COL_AMOUNT]
        };

        if (isAuthorized) {
          // ถ้ารหัสถูก: ให้ข้อมูลครบ
          return {
            ...baseData,
            name: r[COL_NAME],
            gift: r[COL_GIFT],
            notes: r[COL_NOTES]
          };
        } else {
          // ถ้ารหัสผิด/ไม่ใส่: เซ็นเซอร์ข้อมูลส่วนตัว
          return {
            ...baseData,
            name: "🔒 Protected",
            gift: false, // ซ่อนสถานะของขวัญ
            notes: ""    // ซ่อนโน้ต
          };
        }
      });

      return createJSONOutput({ 
        status: 'success', 
        authorized: isAuthorized, // บอก Frontend ว่า Login ผ่านไหม
        data: result 
      });

    } else if (action === 'update' || action === 'add') {
      // --- WRITE DATA (ต้องใช้รหัสผ่านเท่านั้น) ---
      
      if (!isAuthorized) {
        return createJSONOutput({ status: 'error', message: 'Unauthorized: Incorrect password' });
      }

      const data = body.data;

      if (action === 'update' && data.row) {
        const row = parseInt(data.row);
        sheet.getRange(row, COL_CHANNEL + 1).setValue(data.channel);
        sheet.getRange(row, COL_AMOUNT + 1).setValue(data.amount);
        sheet.getRange(row, COL_GIFT + 1).setValue(data.gift); // รับ boolean หรือ text
        sheet.getRange(row, COL_NOTES + 1).setValue(data.notes);
        return createJSONOutput({ status: 'success', message: 'Updated' });

      } else if (action === 'add') {
        sheet.appendRow([
          data.name,
          data.side,
          data.relation,
          data.channel,
          data.amount,
          data.gift,
          data.notes
        ]);
        return createJSONOutput({ status: 'success', message: 'Added' });
      }
    }

  } catch (error) {
    return createJSONOutput({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

function createJSONOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}