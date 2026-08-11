const SHEET_NAME = 'Events';
const LINE_CHANNEL_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN_HERE';
const TARGET_USER_OR_GROUP_ID = 'YOUR_TARGET_ID_HERE';
const ADMIN_PIN = '1234';

function doGet(e) {
  try {
    var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    
    // API สำหรับส่งข้อมูล JSON ให้ Vercel
    if (action === 'getEvents') {
      var data = getCanteenEvents();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // กำหนดชื่อไฟล์ตาม Parameter
    var fileName = (page === 'admin') ? 'Admin' : 'Index';
    var pageTitle = (page === 'admin') 
      ? 'PRTC-CCIS | ระบบจัดการหลังบ้านแอดมิน' 
      : 'PRTC-CCIS | ระบบแจ้งการจัดเลี้ยงโรงอาหาร วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา';

    // ใช้ createTemplateFromFile และ evaluate เพื่อรันส่วนประกอบ HTML ให้สมบูรณ์
    return HtmlService.createTemplateFromFile(fileName)
      .evaluate()
      .setTitle(pageTitle)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    // หากหาไฟล์ไม่เจอ หรือมีสคริปต์ใน HTML พัง ให้พ่น Error ออกหน้าจอแทนการปล่อยหน้าขาว
    return HtmlService.createHtmlOutput(
      '<div style="padding: 30px; font-family: sans-serif; text-align: center; color: #721c24; background-color: #f8d7da; border-radius: 8px; margin: 20px;">' +
      '<h3 style="margin-top: 0;">❌ เกิดข้อผิดพลาดในการโหลดหน้าเว็บ (' + (e.parameter.page || 'index') + ')</h3>' +
      '<p><strong>รายละเอียด:</strong> ' + err.toString() + '</p>' +
      '<p style="font-size: 0.9em; color: #6c757d;">โปรดตรวจสอบว่าฝั่งซ้ายมือมีไฟล์ <b>Index.html</b> และ <b>Admin.html</b> (ตัวพิมพ์ใหญ่-เล็กตรงเป๊ะ) แล้วหรือยัง</p>' +
      '</div>'
    );
  }
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

    try { sendLineNotification(form); } catch(e) {}
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

function sendLineNotification(data) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN.includes('YOUR_')) return;
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
          { "type": "text", "text": `🍱 เมนูอาหาร: ${data.mainMenu}`, "size": "sm", "wrap": true, "weight": "bold", "color": "#1b5e20", "margin": "md" },
          { "type": "text", "text": `🍦 ของหวาน/ของทานเล่น: ${data.snackDessert || '-'}`, "size": "sm", "wrap": true, "color": "#e65100" }
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
