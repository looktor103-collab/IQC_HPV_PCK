// ============================================================
//  BSC Maintenance Form — Google Apps Script
//  วิธีใช้:
//    1. เปิด Google Sheets → Extensions → Apps Script
//    2. ลบโค้ดเดิมทั้งหมด → วางโค้ดนี้
//    3. Deploy → New Deployment → Web app
//       Execute as: Me  |  Who has access: Anyone
//    4. คัดลอก Web app URL → วางใน BSC Form (ปุ่ม ⚙️)
// ============================================================

/**
 * รับ POST จาก BSC Maintenance Form
 * body (text/plain JSON): { rows: string[][], month: string, sheetName: string }
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = payload.sheetName || payload.month || 'BSC';

    // หา tab หรือสร้างใหม่
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
    } else {
      sheet.clearContents();
      sheet.clearFormats();
    }

    // เขียนข้อมูล
    if (payload.rows && payload.rows.length > 0) {
      const numCols = Math.max(...payload.rows.map(r => r.length));
      // ทำให้ทุกแถวมีจำนวน column เท่ากัน
      const normalized = payload.rows.map(r => {
        const row = [...r];
        while (row.length < numCols) row.push('');
        return row;
      });
      sheet.getRange(1, 1, normalized.length, numCols).setValues(normalized);

      // จัดรูปแบบแถวหัว
      sheet.getRange(1, 1, 1, numCols)
           .setBackground('#1d4ed8')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setFontSize(10);

      // แถว column headers (วันที่)
      const headerRowIndex = normalized.findIndex(r => r[0] === 'รายการ') + 1;
      if (headerRowIndex > 0) {
        sheet.getRange(headerRowIndex, 1, 1, numCols)
             .setBackground('#1e40af')
             .setFontColor('#ffffff')
             .setFontWeight('bold');
      }

      // ไฮไลต์ช่องที่มีเครื่องหมาย /
      for (let i = 0; i < normalized.length; i++) {
        for (let j = 1; j < numCols; j++) {
          if (normalized[i][j] === '/') {
            sheet.getRange(i + 1, j + 1).setBackground('#d1fae5').setFontWeight('bold');
          }
        }
      }

      // ไฮไลต์ช่องผู้ตรวจสอบที่เซ็นแล้ว
      const inspRowIndex = normalized.findIndex(r => r[0] === 'ผู้ตรวจสอบ') + 1;
      if (inspRowIndex > 0) {
        for (let j = 1; j < numCols; j++) {
          if (normalized[inspRowIndex - 1][j] === '✓') {
            sheet.getRange(inspRowIndex, j + 1).setBackground('#bfdbfe');
          }
        }
      }

      // ล็อกความกว้างคอลัมน์แรก (รายการ)
      sheet.setColumnWidth(1, 280);
      for (let c = 2; c <= numCols; c++) sheet.setColumnWidth(c, 36);

      // freeze row headers
      if (headerRowIndex > 0) sheet.setFrozenRows(headerRowIndex);
      sheet.setFrozenColumns(1);
    }

    // บันทึก log วันเวลาที่อัปเดต
    const logSheet = ss.getSheetByName('_log') || ss.insertSheet('_log');
    logSheet.appendRow([new Date(), tabName, 'บันทึกสำเร็จ', payload.rows?.length + ' แถว']);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, tab: tabName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** ทดสอบว่า script พร้อมใช้งาน */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'BSC Maintenance Script ready ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}
