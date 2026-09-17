const { google } = require('googleapis');
const path = require('path');

// ==========================================
// CONFIGURATION
// ==========================================
const CONFIG = {
  SPREADSHEET_ID: 'วาง_SPREADSHEET_ID_ตรงนี้', // ใส่ Google Sheet ID ของคุณ
  KEY_FILE_PATH: path.join(__dirname, 'credentials.json'), // ไฟล์ Service Account
  MAIN_CODES_PER_SHOW: 50,
  PREFIX: 'TDLLIVE',
  TIMEZONE: 'Asia/Bangkok'
};

// ==========================================
// AUTHENTICATION & GOOGLE SHEETS CLIENT
// ==========================================
const auth = new google.auth.GoogleAuth({
  keyFile: CONFIG.KEY_FILE_PATH,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

/**
 * ดึงข้อมูลตาม Range จาก Google Sheets
 */
async function getSheetValues(range) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: range,
  });
  return response.data.values || [];
}

/**
 * เพิ่มแถวใหม่ลง Google Sheets
 */
async function appendSheetValues(range, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: CONFIG.SPREADSHEET_ID,
    range: range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

// ==========================================
// SETUP SYSTEM / INITIALIZE SHEETS
// ==========================================
async function setup() {
  try {
    // ล้างข้อมูลและตั้ง Header
    await sheets.spreadsheets.values.clear({ spreadsheetId: CONFIG.SPREADSHEET_ID, range: 'SHOWS!A1:Z' });
    await sheets.spreadsheets.values.clear({ spreadsheetId: CONFIG.SPREADSHEET_ID, range: 'CODES!A1:Z' });
    await sheets.spreadsheets.values.clear({ spreadsheetId: CONFIG.SPREADSHEET_ID, range: 'CHAT!A1:Z' });

    await appendSheetValues('SHOWS!A1', [['SHOW_ID', 'DATE', 'TIME', 'SHOW_NAME', 'POSTER_URL', 'YOUTUBE_ID', 'ACTIVE']]);
    await appendSheetValues('CODES!A1', [['CODE', 'SHOW_ID', 'STATUS', 'CREATED_AT', 'MAX_USES']]);
    await appendSheetValues('CHAT!A1', [['SHOW_ID', 'NICKNAME', 'MESSAGE', 'TIMESTAMP', 'ID']]);

    console.log('✅ สร้างระบบเรียบร้อยแล้ว (SHOWS, CODES, CHAT)');
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการ Setup:', error.message);
  }
}

// ==========================================
// SHOWS & YOUTUBE
// ==========================================
function extractYoutubeId(urlOrId) {
  if (!urlOrId) return '';
  const str = String(urlOrId).trim();
  const match = str.match(/(?:youtu\.be\/|youtube\.com\/(?:embed\/|v\/|watch\?v=|watch\?.+&v=))([\w-]{11})/);
  return match ? match[1] : (str.length === 11 ? str : str);
}

function formatDate(val) {
  if (!val) return '';
  const date = new Date(val);
  if (isNaN(date.getTime())) return String(val);
  
  return new Intl.DateTimeFormat('th-TH', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    timeZone: CONFIG.TIMEZONE
  }).format(date);
}

function formatTime(val) {
  if (!val) return '';
  const date = new Date(val);
  if (isNaN(date.getTime())) return String(val);

  return new Intl.DateTimeFormat('th-TH', {
    hour: '2-digit', minute: '2-digit', hour12: false,
    timeZone: CONFIG.TIMEZONE
  }).format(date);
}

async function getShows() {
  try {
    const values = await getSheetValues('SHOWS!A2:G');
    if (!values.length) return [];

    return values
      .filter(r => r[0] && String(r[6] || '').toUpperCase() !== 'NO')
      .map(r => ({
        id: String(r[0]),
        date: formatDate(r[1]),
        time: formatTime(r[2]),
        name: String(r[3] || ''),
        poster: String(r[4] || ''),
        youtube: extractYoutubeId(r[5]),
        active: String(r[6] || 'YES')
      }));
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการดึงข้อมูล SHOWS:', error.message);
    return [];
  }
}

// ==========================================
// CODE GENERATOR
// ==========================================
function createLongCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let rand = '';
  for (let i = 0; i < 12; i++) {
    rand += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return CONFIG.PREFIX + '-' + rand.substring(0, 4) + '-' + rand.substring(4, 8) + '-' + rand.substring(8, 12);
}

async function generateCodes(showId) {
  const now = new Date().toISOString();
  const rows = [];

  for (let i = 0; i < CONFIG.MAIN_CODES_PER_SHOW; i++) {
    rows.push([createLongCode(), showId, 'ACTIVE', now, 'UNLIMITED']);
  }

  await appendSheetValues('CODES!A:E', rows);
}

async function generateCodesForAllShows() {
  try {
    const shows = await getShows();
    if (!shows.length) {
      console.log('⚠️ ไม่พบข้อมูลรอบการแสดงใน SHOWS');
      return;
    }

    for (const show of shows) {
      await generateCodes(show.id);
    }

    console.log('✅ สร้าง Codes รูปแบบใหม่สำเร็จ!');
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการสร้าง Codes:', error.message);
  }
}

// ==========================================
// VERIFY CODE
// ==========================================
async function verifyCode(showId, code) {
  code = String(code || '').trim().toUpperCase();
  if (!showId || !code) return { success: false, message: 'กรุณากรอกรหัสเข้าชม' };

  try {
    const data = await getSheetValues('CODES!A2:C');
    if (!data.length) return { success: false, message: 'ไม่พบข้อมูลรหัส' };

    for (let i = 0; i < data.length; i++) {
      const savedCode = String(data[i][0] || '').trim().toUpperCase();
      const savedShow = String(data[i][1] || '').trim();
      const status = String(data[i][2] || '').trim().toUpperCase();

      if (savedCode === code && savedShow === String(showId)) {
        if (status !== 'ACTIVE') {
          return { success: false, message: 'รหัสนี้ถูกยกเลิกการใช้งานแล้ว' };
        }
        return { success: true };
      }
    }

    return { success: false, message: 'Access Code ไม่ถูกต้อง' };
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการตรวจสอบ Code:', error.message);
    return { success: false, message: 'เกิดข้อผิดพลาดจากระบบ' };
  }
}

// ==========================================
// REAL-TIME CHAT
// ==========================================
async function sendChatMessage(showId, nickname, message) {
  if (!showId || !message.trim()) return { success: false };
  try {
    const now = new Date();
    const id = 'MSG-' + now.getTime();

    await appendSheetValues('CHAT!A:E', [[
      String(showId),
      String(nickname || 'Anonymous').substring(0, 20),
      String(message).substring(0, 200),
      now.toISOString(),
      id
    ]]);

    return { success: true };
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการส่งข้อความ:', error.message);
    return { success: false };
  }
}

async function getChatMessages(showId) {
  try {
    const values = await getSheetValues('CHAT!A2:D');
    if (!values.length) return [];

    return values
      .filter(r => String(r[0]) === String(showId))
      .slice(-50)
      .map(r => ({
        user: String(r[1]),
        text: String(r[2]),
        time: formatTime(r[3])
      }));
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาดในการดึงข้อความแชท:', error.message);
    return [];
  }
}

// ==========================================
// EXPORTS FOR SERVER (Express / Node.js)
// ==========================================
module.exports = {
  setup,
  getShows,
  generateCodesForAllShows,
  verifyCode,
  sendChatMessage,
  getChatMessages
};
