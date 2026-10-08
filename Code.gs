/**
 * =========================================================================
 * 🩺 Min's Daily Health Log & Stock API
 * Google Apps Script Web App (สำหรับ Google Sheets)
 * =========================================================================
 * 
 * วิธีติดตั้งใน Google Sheets:
 * 1. เปิด Google Sheet ของน้องมิน
 * 2. ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3. วางโค้ดทั้งหมดนี้ลงในไฟล์ Code.gs (แทนที่โค้ดเดิม)
 * 4. กดปุ่ม "ทำให้ใช้งานได้" (Deploy) > จัดการการทำให้ใช้งานได้ (Manage deployments)
 * 5. กดไอคอนรูปดินสอ (Edit) > เลือกเวอร์ชัน "ใหม่" (New version) > บันทึก (Deploy)
 */

// 🔑 Secret Token สำหรับป้องกันการเขียนข้อมูล
const SECRET_KEY = "MinHealthLogSecret2026_SecureKey_XyZ99";

/**
 * =========================================================================
 * 📥 1. ส่วน GET API (doGet): ดึงข้อมูลสต็อกยา นม และอาหารเสริม
 * =========================================================================
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ค้นหาแผ่นงานสต็อก (ตรวจหาชื่อ Stock_Log หรือชื่อใกล้เคียง)
    let sheet = ss.getSheetByName("Stock_Log");
    if (!sheet) sheet = ss.getSheetByName("Stock");
    if (!sheet) sheet = ss.getSheetByName("Stock_Master");

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "ไม่พบชีตชื่อ 'Stock_Log' กรุณาตรวจสอบชื่อแท็บใน Google Sheets"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const rawHeaders = data[0];
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // ข้ามแถวที่ว่างทั้งหมด
      const isRowEmpty = row.every(cell => cell === "" || cell === null || cell === undefined);
      if (isRowEmpty) continue;

      const item = {};
      rawHeaders.forEach((h, colIndex) => {
        const headerName = String(h).trim();
        if (!headerName) return;

        let val = row[colIndex];
        // แปลง Date ในชีตเป็นสตริงมาตรฐาน YYYY-MM-DD
        if (val instanceof Date) {
          val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
        }
        item[headerName] = val;
      });
      rows.push(item);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      count: rows.length,
      updated_at: Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
      data: rows
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * =========================================================================
 * 📤 2. ส่วน POST API (doPost): บันทึกข้อมูลรอบชัก & จัดลำดับอัตโนมัติ
 * =========================================================================
 */
function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const data = JSON.parse(rawData);

    // ตรวจสอบ Secret Token
    if (data.secret !== SECRET_KEY) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Unauthorized: Invalid Secret Token"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // ชีตแรก (ข้อมูลการชัก)
    const sheet = ss.getSheets()[0];

    // เพิ่มแถวข้อมูลใหม่ (14 คอลัมน์)
    // Date, Time, Round, DocCount, Hour, Day, Gap, VNS, VNS Result, Symptom, Duration, Trigger, Emergency Med, Note
    const newRow = [
      data.date,
      data.time_start,
      "", // Round (คำนวณอัตโนมัติ)
      "", // DocCount (คำนวณอัตโนมัติ)
      data.hour !== undefined ? data.hour : "",
      data.day || "",
      "", // Gap
      data.vns || "No",
      data.vns_result || "ไม่ได้ใช้",
      data.symptom_main || "",
      data.duration || 1,
      data.trigger || "",
      data.emergency_med || "",
      data.note || ""
    ];

    sheet.appendRow(newRow);

    // จัดเรียงและคำนวณรอบชัก
    recalculateRounds(sheet);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Recorded and sorted successfully"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 🔄 ฟังก์ชันจัดเรียงตามวันที่+เวลา และคำนวณรอบชักใหม่อัตโนมัติ
 */
function recalculateRounds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // เรียงข้อมูลตามวันที่ (Col 1) และเวลา (Col 2)
  const range = sheet.getRange(2, 1, lastRow - 1, 14);
  range.sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true }
  ]);

  const values = range.getValues();
  let currentDay = "";
  let dayRound = 0;

  for (let i = 0; i < values.length; i++) {
    const rowDate = Utilities.formatDate(new Date(values[i][0]), "GMT+7", "yyyy-MM-dd");
    const symptom = String(values[i][9] || "");

    // ถ้าเปลี่ยนวัน ให้รีเซ็ตรอบชัก
    if (rowDate !== currentDay) {
      currentDay = rowDate;
      dayRound = 0;
    }

    // ถ้าไม่ใช่พบแพทย์ ให้นับเป็นรอบชัก
    if (symptom.indexOf("พบแพทย์") === -1) {
      dayRound++;
      values[i][2] = dayRound; // Col C: รอบชักของวัน
      values[i][3] = dayRound === 1 ? 1 : 0; // Col D: นับวันชัก (หมอ)
    } else {
      values[i][2] = 0;
      values[i][3] = 0;
    }
  }

  // บันทึกค่า Col C และ Col D กลับลงชีต
  sheet.getRange(2, 1, values.length, 14).setValues(values);
}
