// ID ของตาราง Google Sheets
const SPREADSHEET_ID = '1c1KshY9aaRBK-B-SbWA3mVuDJrtlQg9ZmAkB3mZXc2Q'; 
const SHEET_NAME = 'Events';
const LINE_CHANNEL_ACCESS_TOKEN = 'YOUR_LINE_CHANNEL_ACCESS_TOKEN_HERE';
const TARGET_USER_OR_GROUP_ID = 'YOUR_TARGET_ID_HERE';
const ADMIN_PIN = '1234';

function doGet(e) {
  try {
    var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    
    // API สำหรับดึงข้อมูล JSON
    if (action === 'getEvents') {
      var data = getCanteenEvents();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // เลือกรุ่น HTML ออกแสดงผลตามหน้า
    var htmlContent = (page === 'admin') ? getAdminHtml() : getIndexHtml();
    
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle('PRTC-CCIS | วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="padding: 20px; color: red; font-family: sans-serif;">' +
      '<h3>เกิดข้อผิดพลาดในการโหลดระบบ:</h3><p>' + err.toString() + '</p></div>'
    );
  }
}

function checkAdminPin(pin) {
  return String(pin) === String(ADMIN_PIN);
}

function getCanteenEvents() {
  try {
    let sheet;
    if (SPREADSHEET_ID && !SPREADSHEET_ID.includes('ใส่_SPREADSHEET_ID')) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    } else {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    }
    
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
    let ss;
    if (SPREADSHEET_ID && !SPREADSHEET_ID.includes('ใส่_SPREADSHEET_ID')) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

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
  
  let sheet;
  if (SPREADSHEET_ID && !SPREADSHEET_ID.includes('ใส่_SPREADSHEET_ID')) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  } else {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  }

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
    "altText": "[PRTC-CCIS] แจ้งข่าวการจัดเลี้ยง: " + data.occasion,
    "contents": {
      "type": "bubble",
      "header": {
        "type": "box", "layout": "vertical", "backgroundColor": statusColor,
        "contents": [
          { "type": "text", "text": "📢 PRTC Canteen Catering Alert", "color": "#ffffff", "weight": "bold", "size": "sm" },
          { "type": "text", "text": "สถานะ: " + data.guestStatus, "color": "#ffffff", "size": "xs" }
        ]
      },
      "body": {
        "type": "box", "layout": "vertical",
        "contents": [
          { "type": "text", "text": data.occasion, "weight": "bold", "size": "md", "wrap": true },
          { "type": "separator", "margin": "md" },
          { "type": "text", "text": "📅 วันที่: " + data.date + " (" + data.time + ")", "size": "sm", "margin": "md" },
          { "type": "text", "text": "👤 แขก/เจ้าภาพ: " + data.guestName + " (" + data.guestCount + " ท่าน)", "size": "sm" },
          { "type": "text", "text": "🍱 เมนูอาหาร: " + data.mainMenu, "size": "sm", "wrap": true, "weight": "bold", "color": "#1b5e20", "margin": "md" },
          { "type": "text", "text": "🍦 ของหวาน/ของทานเล่น: " + (data.snackDessert || '-'), "size": "sm", "wrap": true, "color": "#e65100" }
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

function getIndexHtml() {
  return '<!DOCTYPE html>' +
'<html lang="th">' +
'<head>' +
'  <base target="_top">' +
'  <meta charset="utf-8">' +
'  <meta name="viewport" content="width=device-width, initial-scale=1">' +
'  <title>PRTC-CCIS | ระบบแจ้งการจัดเลี้ยงโรงอาหาร</title>' +
'  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">' +
'  <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600&display=swap" rel="stylesheet">' +
'  <style>' +
'    body { background-color: #f4f6f9; font-family: "Prompt", sans-serif; }' +
'    .prtc-header { background: linear-gradient(135deg, #004085 0%, #002752 100%); color: #fff; border-radius: 12px; }' +
'    .card-event { border-radius: 12px; border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.06); transition: transform 0.2s; }' +
'    .card-event:hover { transform: translateY(-2px); }' +
'  </style>' +
'</head>' +
'<body class="p-3 p-md-4">' +
'  <div class="container" style="max-width: 850px;">' +
'    <div class="prtc-header p-4 mb-4 shadow-sm text-center">' +
'      <h3 class="fw-bold mb-1">PRTC Canteen Catering Information System</h3>' +
'      <p class="mb-2 text-white-50" style="font-size: 0.9rem;">ระบบสารสนเทศแจ้งการจัดเลี้ยงโรงอาหาร | วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา</p>' +
'      <button onclick="goToAdmin()" class="btn btn-sm btn-outline-light rounded-pill px-3 mt-1">' +
'        ⚙️ เข้าสู่หน้าแอดมินบันทึกข้อมูล' +
'      </button>' +
'    </div>' +
'    <div class="card p-3 border-0 shadow-sm mb-4">' +
'      <div class="row g-2">' +
'        <div class="col-md-5">' +
'          <input type="text" id="searchInput" class="form-control" placeholder="🔍 ค้นหาตามเมนู, ชื่อแขก, โอกาส..." onkeyup="filterEvents()">' +
'        </div>' +
'        <div class="col-md-4">' +
'          <select id="statusFilter" class="form-select" onchange="filterEvents()">' +
'            <option value="ALL">สถานะทั้งหมด</option>' +
'            <option value="มา">เฉพาะยืนยันแล้ว (มา)</option>' +
'            <option value="รอยืนยัน">รอยืนยัน</option>' +
'            <option value="ไม่มา/ยกเลิก">ไม่มา/ยกเลิก</option>' +
'          </select>' +
'        </div>' +
'        <div class="col-md-3">' +
'          <button class="btn btn-primary w-100 rounded-pill" onclick="loadEvents()">🔄 รีเฟรชข้อมูล</button>' +
'        </div>' +
'      </div>' +
'    </div>' +
'    <div class="d-flex justify-content-between align-items-center mb-3">' +
'      <h5 class="fw-bold text-secondary mb-0">📅 รายการจัดเลี้ยงอาหาร</h5>' +
'      <span id="eventCount" class="badge bg-secondary">0 รายการ</span>' +
'    </div>' +
'    <div id="eventsList" class="row g-3">' +
'      <div class="text-center py-5 text-muted">กำลังดึงข้อมูลระบบ PRTC-CCIS...</div>' +
'    </div>' +
'  </div>' +
'  <script>' +
'    let allEvents = [];' +
'    function goToAdmin() {' +
'      const currentUrl = window.location.href.split("?")[0];' +
'      window.location.href = currentUrl + "?page=admin";' +
'    }' +
'    function loadEvents() {' +
'      document.getElementById("eventsList").innerHTML = "<div class=\'text-center py-5 text-muted\'>กำลังดึงข้อมูลระบบ PRTC-CCIS...</div>";' +
'      google.script.run' +
'        .withSuccessHandler(data => {' +
'          allEvents = data || [];' +
'          filterEvents();' +
'        })' +
'        .getCanteenEvents();' +
'    }' +
'    function filterEvents() {' +
'      const searchText = document.getElementById("searchInput").value.toLowerCase();' +
'      const statusValue = document.getElementById("statusFilter").value;' +
'      const filtered = allEvents.filter(e => {' +
'        const matchesSearch = (e.occasion + e.guestName + e.mainMenu + e.snackDessert + e.note).toLowerCase().includes(searchText);' +
'        const matchesStatus = statusValue === "ALL" || e.guestStatus === statusValue;' +
'        return matchesSearch && matchesStatus;' +
'      });' +
'      document.getElementById("eventCount").innerText = filtered.length + " รายการ";' +
'      renderEvents(filtered);' +
'    }' +
'    function renderEvents(events) {' +
'      const container = document.getElementById("eventsList");' +
'      if (!events || events.length === 0) {' +
'        container.innerHTML = "<div class=\'alert alert-light text-center border shadow-sm\'>ไม่พบรายการจัดเลี้ยงอาหาร</div>";' +
'        return;' +
'      }' +
'      container.innerHTML = events.map(e => {' +
'        let badgeColor = "bg-success";' +
'        if (e.guestStatus === "ไม่มา/ยกเลิก") badgeColor = "bg-danger";' +
'        if (e.guestStatus === "รอยืนยัน") badgeColor = "bg-warning text-dark";' +
'        return \'<div class="col-12"><div class="card card-event p-3 p-md-4"><div class="d-flex justify-content-between align-items-center mb-2"><span class="badge \' + badgeColor + \' px-3 py-2 rounded-pill">\' + e.guestStatus + \'</span><small class="text-muted fw-semibold">📅 \' + e.date + \' | ⏰ \' + e.time + \'</small></div><h5 class="fw-bold text-dark mb-2">\' + e.occasion + \'</h5><p class="text-muted mb-2" style="font-size: 0.95rem;"><strong>เจ้าภาพ/แขก:</strong> \' + e.guestName + \' (\' + e.guestCount + \' ท่าน) \' + (e.hasFamily === "ใช่" ? \'<span class="badge bg-info-subtle text-info-emphasis ms-1">👨‍👩‍👧‍👦 มีครอบครัว/ผู้ติดตาม</span>\' : \'\') + \'</p><hr class="my-2"><div class="row g-2 mt-1"><div class="col-md-6"><small class="text-muted d-block">🍱 รายการอาหารหลัก:</small><strong class="text-success">\' + e.mainMenu + \'</strong></div><div class="col-md-6"><small class="text-muted d-block">🍦 ของหวาน / ของทานเล่น:</small><strong class="text-warning-emphasis">\' + (e.snackDessert || \'-\') + \'</strong></div></div>\' + (e.note ? \'<div class="mt-2 text-secondary small">📌 <strong>หมายเหตุ:</strong> \' + e.note + \'</div>\' : \'\') + \'</div></div>\';' +
'      }).join("");' +
'    }' +
'    window.onload = loadEvents;' +
'  </script>' +
'</body>' +
'</html>';
}

function getAdminHtml() {
  return '<!DOCTYPE html>' +
'<html lang="th">' +
'<head>' +
'  <base target="_top">' +
'  <meta charset="utf-8">' +
'  <meta name="viewport" content="width=device-width, initial-scale=1">' +
'  <title>PRTC-CCIS | บันทึกข้อมูลการจัดเลี้ยง (Admin)</title>' +
'  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">' +
'  <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">' +
'  <style>' +
'    body { background-color: #f0f2f5; font-family: "Prompt", sans-serif; color: #333; }' +
'    .admin-header { background: linear-gradient(135deg, #004085 0%, #002752 100%); color: #ffffff; border-radius: 16px; box-shadow: 0 8px 20px rgba(0, 39, 82, 0.15); }' +
'    .card-form { border-radius: 16px; border: none; box-shadow: 0 4px 18px rgba(0,0,0,0.05); }' +
'    .form-label { font-weight: 500; color: #495057; margin-bottom: 0.3rem; }' +
'    .form-control, .form-select { border-radius: 10px; padding: 0.6rem 0.9rem; border: 1px solid #ced4da; }' +
'    .form-control:focus, .form-select:focus { border-color: #004085; box-shadow: 0 0 0 0.25rem rgba(0, 64, 133, 0.15); }' +
'    .btn-submit { background-color: #004085; border: none; border-radius: 10px; padding: 0.75rem; font-weight: 600; font-size: 1rem; transition: all 0.2s; }' +
'    .btn-submit:hover { background-color: #002752; }' +
'    .login-card { max-width: 420px; margin: 80px auto auto; border-radius: 20px; border: none; box-shadow: 0 10px 30px rgba(0,0,0,0.08); }' +
'    .table-card { border-radius: 16px; border: none; box-shadow: 0 4px 18px rgba(0,0,0,0.05); }' +
'  </style>' +
'</head>' +
'<body class="p-3 p-md-4">' +
'  <div class="container" style="max-width: 800px;">' +
'    <div id="authScreen" class="card login-card p-4 p-md-5 bg-white text-center">' +
'      <div class="mb-3 fs-1 text-primary">🔑</div>' +
'      <h4 class="fw-bold mb-2">PRTC-CCIS Admin</h4>' +
'      <p class="text-muted small mb-4">เข้าสู่ระบบสำหรับแอดมิน กรอกรหัส PIN เพื่อลงข้อมูล</p>' +
'      <div class="mb-3">' +
'        <input type="password" id="pinInput" class="form-control text-center fs-3 fw-bold" maxlength="6" placeholder="••••" autofocus>' +
'      </div>' +
'      <button class="btn btn-primary btn-submit w-100 mb-2" onclick="loginAdmin()">เข้าสู่ระบบ</button>' +
'      <a href="./" class="btn btn-link text-decoration-none text-muted small">← กลับไปหน้าตารางอาหารสำหรับคนในวิทยาลัย</a>' +
'    </div>' +
'    <div id="adminPanel" style="display: none;">' +
'      <div class="admin-header p-4 mb-4 d-flex justify-content-between align-items-center">' +
'        <div>' +
'          <h4 class="fw-bold mb-1">⚙️ ฟอร์มบันทึกข้อมูลการจัดเลี้ยง</h4>' +
'          <p class="mb-0 text-white-50 small">วิทยาลัยเทคโนโลยีพระมหาไถ่ พัทยา (PRTC-CCIS)</p>' +
'        </div>' +
'        <div>' +
'          <a href="./" class="btn btn-outline-light btn-sm rounded-pill px-3 me-2">หน้าตารางอาหาร</a>' +
'          <button class="btn btn-outline-danger btn-sm rounded-pill px-3" onclick="logoutAdmin()">ออกจากระบบ</button>' +
'        </div>' +
'      </div>' +
'      <div class="card card-form p-4 mb-4 bg-white">' +
'        <div class="d-flex justify-content-between align-items-center mb-3">' +
'          <h5 id="formTitle" class="fw-bold text-primary mb-0">➕ เพิ่มรายการจัดเลี้ยงใหม่</h5>' +
'          <span class="badge bg-light text-dark fw-normal">แอดมินส่วนกลาง</span>' +
'        </div>' +
'        <form id="adminForm" onsubmit="handleFormSubmit(event)">' +
'          <input type="hidden" id="id" name="id">' +
'          <div class="row g-3 mb-3">' +
'            <div class="col-md-6">' +
'              <label class="form-label">📅 วันที่จัดเลี้ยง <span class="text-danger">*</span></label>' +
'              <input type="date" class="form-control" id="date" name="date" required>' +
'            </div>' +
'            <div class="col-md-6">' +
'              <label class="form-label">⏰ เวลาช่วงที่จัด <span class="text-danger">*</span></label>' +
'              <input type="text" class="form-control" id="time" name="time" placeholder="เช่น 11:30 - 13:00 น." required>' +
'            </div>' +
'          </div>' +
'          <div class="mb-3">' +
'            <label class="form-label">🎉 เนื่องในโอกาส / งาน <span class="text-danger">*</span></label>' +
'            <input type="text" class="form-control" id="occasion" name="occasion" placeholder="เช่น เลี้ยงวันเกิด / คณะศึกษาดูงานบริจาคทุน" required>' +
'          </div>' +
'          <div class="row g-3 mb-3">' +
'            <div class="col-md-8">' +
'              <label class="form-label">👤 ชื่อแขกผู้มีเกียรติ / เจ้าภาพ <span class="text-danger">*</span></label>' +
'              <input type="text" class="form-control" id="guestName" name="guestName" placeholder="เช่น คุณสมชาย ใจดี และคณะ" required>' +
'            </div>' +
'            <div class="col-md-4">' +
'              <label class="form-label">👥 จำนวนแขก (คน) <span class="text-danger">*</span></label>' +
'              <input type="number" class="form-control" id="guestCount" name="guestCount" min="1" placeholder="เช่น 10" required>' +
'            </div>' +
'          </div>' +
'          <div class="mb-3 p-3 bg-light rounded-3">' +
'            <div class="form-check">' +
'              <input type="checkbox" class="form-check-input" id="hasFamily" name="hasFamily">' +
'              <label class="form-check-label fw-medium" for="hasFamily">👨‍👩‍👧‍👦 มีครอบครัว / ผู้ติดตามมาด้วย</label>' +
'            </div>' +
'          </div>' +
'          <div class="mb-3">' +
'            <label class="form-label">🍱 รายการอาหารหลัก <span class="text-danger">*</span></label>' +
'            <textarea class="form-control" id="mainMenu" name="mainMenu" rows="2" placeholder="เช่น ก๋วยเตี๋ยวเรือหมู, ข้าวมันไก่" required></textarea>' +
'          </div>' +
'          <div class="mb-3">' +
'            <label class="form-label">🍦 ของกินเล่น / ของหวาน / ไอศกรีม</label>' +
'            <input type="text" class="form-control" id="snackDessert" name="snackDessert" placeholder="เช่น ไอศกรีมกะทิสด, ผลไม้ตามฤดูกาล">' +
'          </div>' +
'          <div class="mb-3">' +
'            <label class="form-label">📌 สถานะการยืนยันของแขก <span class="text-danger">*</span></label>' +
'            <select class="form-select fw-medium" id="guestStatus" name="guestStatus">' +
'              <option value="มา" selected class="text-success">✅ มา (ยืนยันแล้ว)</option>' +
'              <option value="รอยืนยัน" class="text-warning">⏳ รอยืนยัน</option>' +
'              <option value="ไม่มา/ยกเลิก" class="text-danger">❌ ไม่มา / ยกเลิก</option>' +
'            </select>' +
'          </div>' +
'          <div class="mb-4">' +
'            <label class="form-label">📝 หมายเหตุเพิ่มเติม</label>' +
'            <input type="text" class="form-control" id="note" name="note" placeholder="เช่น ตั้งโต๊ะชั้น 1 / แจกเฉพาะนักเรียนแผนกช่าง">' +
'          </div>' +
'          <div class="d-flex gap-2">' +
'            <button type="submit" id="btnSubmit" class="btn btn-primary btn-submit flex-grow-1">💾 บันทึกข้อมูล & ส่ง LINE Alert</button>' +
'            <button type="button" id="btnReset" class="btn btn-outline-secondary px-4 rounded-3" onclick="resetAdminForm()">ล้างฟอร์ม</button>' +
'          </div>' +
'        </form>' +
'      </div>' +
'      <div class="card table-card p-4 bg-white">' +
'        <div class="d-flex justify-content-between align-items-center mb-3">' +
'          <h5 class="fw-bold text-dark mb-0">📋 รายการที่บันทึกไว้แล้ว</h5>' +
'          <button class="btn btn-sm btn-outline-primary" onclick="loadAdminEvents()">🔄 ดึงข้อมูลล่าสุด</button>' +
'        </div>' +
'        <div class="table-responsive">' +
'          <table class="table table-hover align-middle mb-0">' +
'            <thead class="table-light">' +
'              <tr>' +
'                <th style="width: 20%;">วัน-เวลา</th>' +
'                <th style="width: 25%;">โอกาส/งาน</th>' +
'                <th style="width: 25%;">เจ้าภาพ/แขก</th>' +
'                <th style="width: 15%;">สถานะ</th>' +
'                <th style="width: 15%; text-align: center;">จัดการ</th>' +
'              </tr>' +
'            </thead>' +
'            <tbody id="adminTableBody">' +
'              <tr><td colspan="5" class="text-center py-4 text-muted">กำลังดึงข้อมูล...</td></tr>' +
'            </tbody>' +
'          </table>' +
'        </div>' +
'      </div>' +
'    </div>' +
'  </div>' +
'  <script>' +
'    let currentPin = "";' +
'    let adminEvents = [];' +
'    function loginAdmin() {' +
'      const pin = document.getElementById("pinInput").value;' +
'      if (!pin) return alert("กรุณากรอกรหัส PIN");' +
'      google.script.run' +
'        .withSuccessHandler(isValid => {' +
'          if (isValid) {' +
'            currentPin = pin;' +
'            document.getElementById("authScreen").style.display = "none";' +
'            document.getElementById("adminPanel").style.display = "block";' +
'            document.getElementById("date").value = new Date().toISOString().split("T")[0];' +
'            loadAdminEvents();' +
'          } else {' +
'            alert("รหัส PIN ไม่ถูกต้อง!");' +
'          }' +
'        })' +
'        .checkAdminPin(pin);' +
'    }' +
'    function logoutAdmin() {' +
'      currentPin = "";' +
'      document.getElementById("pinInput").value = "";' +
'      document.getElementById("authScreen").style.display = "block";' +
'      document.getElementById("adminPanel").style.display = "none";' +
'    }' +
'    function loadAdminEvents() {' +
'      google.script.run' +
'        .withSuccessHandler(data => {' +
'          adminEvents = data || [];' +
'          renderAdminTable(adminEvents);' +
'        })' +
'        .getCanteenEvents();' +
'    }' +
'    function renderAdminTable(events) {' +
'      const tbody = document.getElementById("adminTableBody");' +
'      if (!events || events.length === 0) {' +
'        tbody.innerHTML = "<tr><td colspan=\'5\' class=\'text-center text-muted py-4\'>ไม่มีข้อมูลการจัดเลี้ยงในระบบ</td></tr>";' +
'        return;' +
'      }' +
'      tbody.innerHTML = events.map(e => {' +
'        return "<tr><td><small class=\'fw-bold\'>" + e.date + "</small><br><small class=\'text-muted\'>" + e.time + "</small></td><td><strong>" + e.occasion + "</strong></td><td><small>" + e.guestName + " (" + e.guestCount + " คน)</small></td><td><span class=\'badge " + (e.guestStatus === "มา" ? "bg-success" : e.guestStatus === "รอยืนยัน" ? "bg-warning text-dark" : "bg-danger") + "\'>" + e.guestStatus + "</span></td><td style=\'text-align: center;\'><button class=\'btn btn-sm btn-outline-primary me-1\' onclick=\'editEvent(\"" + e.id + "\")\' title=\'แก้ไข\'>✏️</button><button class=\'btn btn-sm btn-outline-danger\' onclick=\'deleteEvent(\"" + e.id + "\")\' title=\'ลบ\'>🗑️</button></td></tr>";' +
'      }).join("");' +
'    }' +
'    function editEvent(id) {' +
'      const item = adminEvents.find(e => String(e.id) === String(id));' +
'      if (!item) return;' +
'      document.getElementById("id").value = item.id;' +
'      document.getElementById("date").value = item.date;' +
'      document.getElementById("time").value = item.time;' +
'      document.getElementById("occasion").value = item.occasion;' +
'      document.getElementById("guestName").value = item.guestName;' +
'      document.getElementById("guestCount").value = item.guestCount;' +
'      document.getElementById("hasFamily").checked = item.hasFamily === "ใช่";' +
'      document.getElementById("mainMenu").value = item.mainMenu;' +
'      document.getElementById("snackDessert").value = item.snackDessert;' +
'      document.getElementById("guestStatus").value = item.guestStatus;' +
'      document.getElementById("note").value = item.note || "";' +
'      document.getElementById("formTitle").innerText = "✏️ แก้ไขข้อมูลการจัดเลี้ยง";' +
'      window.scrollTo({ top: 0, behavior: "smooth" });' +
'    }' +
'    function resetAdminForm() {' +
'      document.getElementById("adminForm").reset();' +
'      document.getElementById("id").value = "";' +
'      document.getElementById("date").value = new Date().toISOString().split("T")[0];' +
'      document.getElementById("formTitle").innerText = "➕ เพิ่มรายการจัดเลี้ยงใหม่";' +
'    }' +
'    function handleFormSubmit(e) {' +
'      e.preventDefault();' +
'      const btn = document.getElementById("btnSubmit");' +
'      btn.disabled = true;' +
'      btn.innerText = "กำลังบันทึกข้อมูล...";' +
'      const formData = {' +
'        id: document.getElementById("id").value,' +
'        date: document.getElementById("date").value,' +
'        time: document.getElementById("time").value,' +
'        occasion: document.getElementById("occasion").value,' +
'        guestName: document.getElementById("guestName").value,' +
'        guestCount: document.getElementById("guestCount").value,' +
'        hasFamily: document.getElementById("hasFamily").checked,' +
'        mainMenu: document.getElementById("mainMenu").value,' +
'        snackDessert: document.getElementById("snackDessert").value,' +
'        guestStatus: document.getElementById("guestStatus").value,' +
'        note: document.getElementById("note").value' +
'      };' +
'      google.script.run' +
'        .withSuccessHandler(res => {' +
'          alert(res.message);' +
'          if (res.success) {' +
'            resetAdminForm();' +
'            loadAdminEvents();' +
'          }' +
'          btn.disabled = false;' +
'          btn.innerText = "💾 บันทึกข้อมูล & ส่ง LINE Alert";' +
'        })' +
'        .withFailureHandler(err => {' +
'          alert("เกิดข้อผิดพลาด: " + err.toString());' +
'          btn.disabled = false;' +
'          btn.innerText = "💾 บันทึกข้อมูล & ส่ง LINE Alert";' +
'        })' +
'        .saveCanteenEvent(formData, currentPin);' +
'    }' +
'    function deleteEvent(id) {' +
'      if (!confirm("คุณต้องการลบรายการจัดเลี้ยงนี้ใช่หรือไม่?")) return;' +
'      google.script.run' +
'        .withSuccessHandler(res => {' +
'          alert(res.message);' +
'          if (res.success) {' +
'            loadAdminEvents();' +
'          }' +
'        })' +
'        .deleteCanteenEvent(id, currentPin);' +
'    }' +
'  </script>' +
'</body>' +
'</html>';
}
