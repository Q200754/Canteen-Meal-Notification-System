const SPREADSHEET_ID = '1c1KshY9aaRBK-B-SbWA3mVuDJrtlQg9ZmAkB3mZXc2Q'; 
const SHEET_NAME = 'Events';
const LINE_CHANNEL_ACCESS_TOKEN = 'VCTIfa5ni3lQTTKk3GqgvBwEhGV73iHAj5sSwd60qpjibP5JepzMDteKh6mu68fLXqQCDEnZa6Urntz0Dr4Oel7hituvdMEch2oCqAz9fBpqAiBaMvfTeBRtZu/omEAzekBC5OCnCUcu1EhZlM1YrAdB04t89/1O/w1cDnyilFU=';
const TARGET_USER_OR_GROUP_ID = 'Uffe1be9bbe74df5dc7b9091612420bff';

// กำหนด Username และ Password สำหรับแอดมิน
const ADMIN_USERNAME = 'piangfah.admin';
const ADMIN_PASSWORD = 'Prtc@2026';

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  
  if (action === 'getEvents') {
    var data = getCanteenEvents();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var fileName = (page === 'admin') ? 'Admin' : 'Index';
  var pageTitle = (page === 'admin') 
    ? 'PRTC-CCIS | ระบบจัดการหลังบ้านแอดมิน' 
    : 'PRTC-CCIS | ระบบแจ้งการจัดเลี้ยงโรงอาหาร วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา';

  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ตรวจสอบ Username และ Password
function checkAdminAuth(username, password) {
  return String(username) === ADMIN_USERNAME && String(password) === ADMIN_PASSWORD;
}

function getCanteenEvents() {
  try {
    let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rows = data.slice(1);
    return rows.map((row, index) => {
      // แปลงรูปแบบวันที่ปลอดภัย
      let dateStr = '';
      if (row[1]) {
        try {
          dateStr = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch(err) {
          dateStr = String(row[1]);
        }
      }

      // แปลงรูปแบบเวลาปลอดภัย
      let timeStr = '';
      if (row[2]) {
        if (row[2] instanceof Date) {
          timeStr = Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), 'HH:mm');
        } else {
          timeStr = String(row[2]);
        }
      }

      return {
        rowIndex: index + 2,
        id: String(row[0] || ''),
        date: dateStr,
        time: timeStr,
        occasion: String(row[3] || ''),
        guestName: String(row[4] || ''),
        hasFamily: String(row[5] || ''),
        guestCount: row[6] || 0,
        mainMenu: String(row[7] || ''),
        snackDessert: String(row[8] || ''),
        guestStatus: String(row[9] || 'มา'),
        note: row[10] ? String(row[10]) : ''
      };
    });
  } catch (e) {
    return [];
  }
}

function saveCanteenEvent(form, username, password) {
  if (!checkAdminAuth(username, password)) {
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
  try {
    let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
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
      form.hasFamily ? 'มี' : 'ไม่มี', form.guestCount,
      form.mainMenu, form.snackDessert, form.guestStatus, form.note || ''
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    // ส่งการแจ้งเตือน LINE Alert เมื่อบันทึกสำเร็จ
    try { sendLineNotification(form); } catch(e) {}

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.toString() };
  }
}

function deleteCanteenEvent(id, username, password) {
  if (!checkAdminAuth(username, password)) return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
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

// ฟังก์ชันส่งการแจ้งเตือน Flex Message เข้า LINE
function sendLineNotification(data) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !TARGET_USER_OR_GROUP_ID) return;
  
  let statusColor = '#28a745';
  if (data.guestStatus === 'ไม่มา/ยกเลิก') statusColor = '#dc3545';
  if (data.guestStatus === 'รอยืนยัน') statusColor = '#ffc107';

  const flexMessage = {
    "type": "flex",
    "altText": `[PRTC-CCIS] แจ้งข่าวการจัดเลี้ยง: ${data.occasion}`,
    "contents": {
      "type": "bubble",
      "header": {
        "type": "box", "layout": "vertical", "backgroundColor": statusColor,
        "contents": [
          { "type": "text", "text": "📢 PRTC Canteen Catering Alert", "color": "#ffffff", "weight": "bold", "size": "sm" },
          { "type": "text", "text": `สถานะ: ${data.guestStatus}`, "color": "#ffffff", "size": "xs" }
        ]
      },
      "body": {
        "type": "box", "layout": "vertical",
        "contents": [
          { "type": "text", "text": data.occasion, "weight": "bold", "size": "md", "wrap": true },
          { "type": "separator", "margin": "md" },
          { "type": "text", "text": `📅 วันที่: ${data.date} (${data.time})`, "size": "sm", "margin": "md" },
          { "type": "text", "text": `👤 แขก/เจ้าภาพ: ${data.guestName} (${data.guestCount} ท่าน)`, "size": "sm" },
          { "type": "text", "text": `👨‍👩‍👧‍👦 ผู้ติดตาม: ${data.hasFamily ? 'มี' : 'ไม่มี'}`, "size": "sm" },
          { "type": "text", "text": `🍱 เมนูอาหาร: ${data.mainMenu}`, "size": "sm", "wrap": true, "weight": "bold", "color": "#1b5e20", "margin": "md" },
          { "type": "text", "text": `🍦 ของหวาน/ทานเล่น: ${data.snackDessert || '-'}`, "size": "sm", "wrap": true, "color": "#e65100" }
        ]
      }
    }
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    'payload': JSON.stringify({ "to": TARGET_USER_OR_GROUP_ID, "messages": [flexMessage] }),
    'muteHttpExceptions': true
  });
}
