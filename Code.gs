const SHEET_NAME = 'Events';
const LINE_CHANNEL_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN_HERE';
const TARGET_USER_OR_GROUP_ID = 'YOUR_TARGET_ID_HERE';
const ADMIN_PIN = '1234';

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  
  if (action === 'getEvents') {
    var data = getCanteenEvents();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var fileName = (page === 'admin') ? 'Admin' : 'Index';
  
  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle('PRTC-CCIS | วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkAdminPin(pin) {
  return pin === ADMIN_PIN;
}

function getCanteenEvents() {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rows = data.slice(1);
    return rows.map((row, index) => {
      return {
        rowIndex: index + 2,
        id: String(row[0]),
        date: row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        time: row[2],
        occasion: row[3],
        guestName: row[4],
        hasFamily: row[5],
        guestCount: row[6],
        mainMenu: row[7],
        snackDessert: row[8],
        guestStatus: row[9],
        note: row[10]
      };
    });
  } catch (e) {
    return [];
  }
}

function saveCanteenEvent(form, pin) {
  if (!checkAdminPin(pin)) {
    return { success: false, message: 'รหัสผ่าน PIN แอดมินไม่ถูกต้อง' };
  }

  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
      sheet.appendRow(['ID', 'Date', 'Time', 'Occasion', 'GuestName', 'HasFamily', 'GuestCount', 'MainMenu', 'SnackDessert', 'GuestStatus', 'Note']);
    }

    const id = form.id || 'PRTC-EVT-' + new Date().getTime();
    let targetRow = -1;
    if (form.id) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(form.id)) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const rowData = [
      id, form.date, form.time, form.occasion, form.guestName,
      form.hasFamily ? 'ใช่' : 'ไม่ใช่', form.guestCount,
      form.mainMenu, form.snackDessert, form.guestStatus, form.note || ''
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function deleteCanteenEvent(id, pin) {
  if (!checkAdminPin(pin)) return { success: false, message: 'รหัสผ่าน PIN แอดมินไม่ถูกต้อง' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, message: 'ไม่พบแผ่นงาน' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบรายการเรียบร้อยแล้ว' };
    }
  }
  return { success: false, message: 'ไม่พบรายการที่ต้องการลบ' };
}
