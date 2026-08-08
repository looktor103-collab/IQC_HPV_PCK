// วางโค้ดนี้ใน Google Sheet F-CP-68030 → ส่วนเสริม → Apps Script
// Deploy → New deployment → Web App → Execute as: Me, Access: Anyone
// แล้วคัดลอก Web app URL มาใส่ในฟอร์ม

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ชื่อ Sheet tab = ชื่อเดือน เช่น "ส.ค. 2569"
    var sheetName = data.sheetLabel;

    // หา Sheet ที่มีอยู่ หรือสร้างใหม่
    var sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
    } else {
      sh.clearContents();
      sh.clearFormats();
    }

    // เคลียร์เซลล์ที่ merge ไว้ทั้งหมดก่อนเขียนข้อมูล (ต้องทำก่อนเขียนค่า/สี เสมอ
    // ไม่งั้น Sheets จะ error "ตรึงคอลัมน์ที่มีเฉพาะบางส่วนของเซลล์ที่ผสาน")
    var maxR = Math.max(sh.getMaxRows(), 60);
    var maxC = Math.max(sh.getMaxColumns(), 32);
    sh.getRange(1, 1, maxR, maxC).breakApart();

    // แถวที่ 1: หัวข้อ
    sh.getRange(1, 1).setValue('F-CP-68030  แบบบันทึกตารางการบำรุงรักษาเครื่อง cobas 6800 V2.0 system');
    sh.getRange(1, 1, 1, 32).setBackground('#f8fafc').setFontWeight('bold');

    // แถวที่ 2: Serial No และเดือน
    sh.getRange(2, 1).setValue('Serial NO: ' + (data.serialNo || ''));
    sh.getRange(2, 15).setValue('ประจำเดือน: ' + sheetName);

    // แถวที่ 3: Header คอลัมน์ รายการ | 1 | 2 | ... | 31
    var headerRow = ['รายการ'];
    for (var d = 1; d <= 31; d++) { headerRow.push(d); }
    sh.getRange(3, 1, 1, 32).setValues([headerRow]);

    // จัดรูปแบบ header
    var headerRange = sh.getRange(3, 1, 1, 32);
    headerRange.setBackground('#1a3c6e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // แถวข้อมูล tasks
    var tasks = data.tasks;
    var nDays = data.days.length;
    var rowNum = 4;

    for (var i = 0; i < tasks.length; i++) {
      var row = [tasks[i].label];
      for (var d2 = 1; d2 <= 31; d2++) {
        if (d2 <= nDays) {
          row.push(data.checks[tasks[i].key + '_' + d2] || '');
        } else {
          row.push('');  // วันที่ไม่มีในเดือนนั้น
        }
      }
      sh.getRange(rowNum, 1, 1, 32).setValues([row]);

      // เซลล์ที่มี / ทำสีเขียว
      for (var d3 = 1; d3 <= nDays; d3++) {
        if (data.checks[tasks[i].key + '_' + d3] === '/') {
          sh.getRange(rowNum, d3 + 1).setBackground('#ceead6');
        }
      }
      // วันที่เกินเดือนทำสีเทา
      if (nDays < 31) {
        sh.getRange(rowNum, nDays + 2, 1, 31 - nDays).setBackground('#efefef');
      }
      rowNum++;
    }

    // Perform by row (ว่าง)
    var performRow = ['Perform by (ลายเซ็น)'];
    for (var d4 = 1; d4 <= 31; d4++) { performRow.push(''); }
    sh.getRange(rowNum, 1, 1, 32).setValues([performRow]);
    sh.getRange(rowNum, 1, 1, 32).setBackground('#fef7e0');
    sh.getRange(rowNum, 1).setFontWeight('bold');
    rowNum += 2;

    // Additional Note
    sh.getRange(rowNum, 1).setValue('Additional Note :');
    sh.getRange(rowNum, 1).setFontWeight('bold');
    rowNum++;
    sh.getRange(rowNum, 1, 1, 2).setValues([['#', 'บันทึกเพิ่มเติม']]);
    sh.getRange(rowNum, 1, 1, 2).setBackground('#e8f0fe').setFontWeight('bold');
    rowNum++;

    var notes = data.notes || [];
    for (var k = 0; k < notes.length; k++) {
      if (notes[k].text) {
        sh.getRange(rowNum, 1, 1, 2).setValues([[notes[k].no, notes[k].text]]);
        rowNum++;
      }
    }

    // ปรับความกว้างคอลัมน์
    sh.setColumnWidth(1, 280);
    for (var c = 2; c <= 32; c++) { sh.setColumnWidth(c, 32); }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}