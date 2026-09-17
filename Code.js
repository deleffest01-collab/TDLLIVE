const CONFIG = {
  SPREADSHEET_ID: '18_W6uls5wfc9BxLY46zEVPEEJiHW7fzZn2AI__TZux4', // ก๊อปปี้ ID จาก URL Google Sheets มาวางตรงนี้
  MAIN_CODES_PER_SHOW: 50,
  PREFIX: 'TDLLIVE'
};

function getSS() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ==============================
   SETUP & MENU
================================ */
function setup() {
  const ss = getSS();

  let shows = ss.getSheetByName('SHOWS');
  let codes = ss.getSheetByName('CODES');
  let chat = ss.getSheetByName('CHAT');

  if (!shows) shows = ss.insertSheet('SHOWS');
  if (!codes) codes = ss.insertSheet('CODES');
  if (!chat) chat = ss.insertSheet('CHAT');

  shows.clear();
  codes.clear();
  chat.clear();

  shows.getRange(1, 1, 1, 7).setValues([[
    'SHOW_ID', 'DATE', 'TIME', 'SHOW_NAME', 'POSTER_URL', 'YOUTUBE_ID', 'ACTIVE'
  ]]);

  codes.getRange(1, 1, 1, 5).setValues([[
    'CODE', 'SHOW_ID', 'STATUS', 'CREATED_AT', 'MAX_USES'
  ]]);

  chat.getRange(1, 1, 1, 5).setValues([[
    'SHOW_ID', 'NICKNAME', 'MESSAGE', 'TIMESTAMP', 'ID'
  ]]);

  shows.setFrozenRows(1);
  codes.setFrozenRows(1);
  chat.setFrozenRows(1);

  try {
    SpreadsheetApp.getUi().alert('สร้างระบบเรียบร้อยแล้ว (SHOWS, CODES, CHAT)');
  } catch (e) {
    Logger.log('Setup completed successfully via Web App context');
  }
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🎥 PRIVATE LIVE')
      .addItem('⚙️ สร้างระบบ / รีเซ็ต Sheet', 'setup')
      .addItem('🔑 สร้าง Codes (12 หลัก) ให้ทุกรอบ', 'generateCodesForAllShows')
      .addToUi();
  } catch (e) {
    Logger.log('Menu cannot be created in this context');
  }
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('TDL LIVE I แพลตฟอร์มสตรีมมิ่งไลฟ์สดสำหรับงานอีเว้นต์ Roblox')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ==============================
   SHOWS & YOUTUBE
================================ */
function extractYoutubeId(urlOrId) {
  if (!urlOrId) return '';
  const str = String(urlOrId).trim();
  const match = str.match(/(?:youtu\.be\/|youtube\.com\/(?:embed\/|v\/|watch\?v=|watch\?.+&v=))([\w-]{11})/);
  return match ? match[1] : (str.length === 11 ? str : str);
}

function getShows() {
  const sheet = getSS().getSheetByName('SHOWS');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return values
    .filter(r => r[0] && String(r[6]).toUpperCase() !== 'NO')
    .map(r => ({
      id: String(r[0]),
      date: formatDate(r[1]),
      time: formatTime(r[2]),
      name: String(r[3] || ''),
      poster: String(r[4] || ''),
      youtube: extractYoutubeId(r[5]),
      active: String(r[6] || 'YES')
    }));
}

function formatDate(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(val);
}

function formatTime(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(val);
}

/* ==============================
   CODE GENERATOR
================================ */
function generateCodesForAllShows() {
  const shows = getShows();
  if (!shows.length) {
    try {
      SpreadsheetApp.getUi().alert('ยังไม่มีรอบการแสดงใน SHOWS');
    } catch (e) {
      Logger.log('No shows found');
    }
    return;
  }

  shows.forEach(show => generateCodes(show.id));

  try {
    SpreadsheetApp.getUi().alert('สร้าง Codes รูปแบบใหม่สำเร็จ!');
  } catch (e) {
    Logger.log('Codes generated successfully via Web App context');
  }
}

function generateCodes(showId) {
  const sheet = getSS().getSheetByName('CODES');
  const now = new Date();
  const rows = [];

  for (let i = 0; i < CONFIG.MAIN_CODES_PER_SHOW; i++) {
    rows.push([createLongCode(), showId, 'ACTIVE', now, 'UNLIMITED']);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
}

function createLongCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let rand = '';
  for (let i = 0; i < 12; i++) {
    rand += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return CONFIG.PREFIX + '-' + rand.substring(0, 4) + '-' + rand.substring(4, 8) + '-' + rand.substring(8, 12);
}

/* ==============================
   VERIFY CODE
================================ */
function verifyCode(showId, code) {
  code = String(code || '').trim().toUpperCase();
  if (!showId || !code) return { success: false, message: 'กรุณากรอกรหัสเข้าชม' };

  const sheet = getSS().getSheetByName('CODES');
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'ไม่พบข้อมูลรหัส' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  for (let i = 0; i < data.length; i++) {
    const savedCode = String(data[i][0]).trim().toUpperCase();
    const savedShow = String(data[i][1]).trim();
    const status = String(data[i][2]).trim().toUpperCase();

    if (savedCode === code && savedShow === String(showId)) {
      if (status !== 'ACTIVE') {
        return { success: false, message: 'รหัสนี้ถูกยกเลิกการใช้งานแล้ว' };
      }
      return { success: true };
    }
  }

  return { success: false, message: 'Access Code ไม่ถูกต้อง' };
}

/* ==============================
   REAL-TIME CHAT
================================ */
function sendChatMessage(showId, nickname, message) {
  if (!showId || !message.trim()) return { success: false };
  const sheet = getSS().getSheetByName('CHAT');
  const now = new Date();
  const id = 'MSG-' + now.getTime();
  
  sheet.appendRow([
    String(showId),
    String(nickname || 'Anonymous').substring(0, 20),
    String(message).substring(0, 200),
    now,
    id
  ]);
  return { success: true };
}

function getChatMessages(showId) {
  const sheet = getSS().getSheetByName('CHAT');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return values
    .filter(r => String(r[0]) === String(showId))
    .slice(-50)
    .map(r => ({
      user: String(r[1]),
      text: String(r[2]),
      time: Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'HH:mm')
    }));
}