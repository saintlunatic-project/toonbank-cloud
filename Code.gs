/**
 * Toonbank Cloud POS - Backend Google Apps Script
 * Database disimpan dalam sheet yang dipisah sesuai data (SETTINGS, PRODUCTS, dll).
 *
 * Cara pakai:
 * 1) Jika project Apps Script dibuat dari Google Sheets, biarkan SPREADSHEET_ID kosong.
 * 2) Jika project Apps Script standalone, isi SPREADSHEET_ID dengan ID Google Sheet Anda.
 */

const SPREADSHEET_ID = ''; // Opsional: isi jika Apps Script tidak dibuat dari Google Sheet.
const SESSION_TTL_SECONDS = 21600; // 6 jam
const ALLOWED_KEYS = ['SETTINGS', 'PRODUCTS', 'WALLETS', 'TRANSACTIONS', 'HUTANG', 'RIWAYAT_HUTANG', 'MUTASI'];

function doGet(e) {
  setupDatabase();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Toonbank - Premium POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Spreadsheet tidak ditemukan. Isi SPREADSHEET_ID di Code.gs atau buat project Apps Script dari Google Sheet.');
  }
  return active;
}

function getDefaultSettings_() {
  return { idLogin: 'lunatic', idDemo: 'demo', namaToko: 'LUNACELL', namaApp: 'PREMIUM POS' };
}

function setupDatabase() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSpreadsheet_();
    
    const defaults = {
      SETTINGS: JSON.stringify(getDefaultSettings_()),
      PRODUCTS: '[]',
      WALLETS: '[]',
      TRANSACTIONS: '[]',
      HUTANG: '[]',
      RIWAYAT_HUTANG: '[]',
      MUTASI: '[]'
    };

    // Membuat sheet terpisah untuk setiap key data
    ALLOWED_KEYS.forEach(function(key) {
      let sheet = ss.getSheetByName(key);
      if (!sheet) {
        sheet = ss.insertSheet(key);
        sheet.getRange('A1').setValue('Value JSON').setFontWeight('bold');
        sheet.getRange('A2').setValue(defaults[key]);
      }
    });

    return true;
  } finally {
    lock.releaseLock();
  }
}

function readDatabase_() {
  setupDatabase(); // Pastikan semua sheet sudah tersedia
  const ss = getSpreadsheet_();
  const result = {};

  // Membaca data dari masing-masing sheet
  ALLOWED_KEYS.forEach(function(key) {
    const sheet = ss.getSheetByName(key);
    if (sheet) {
      const data = sheet.getRange('A2').getValue();
      try {
        result[key] = JSON.parse(data || (key === 'SETTINGS' ? '{}' : '[]'));
      } catch (err) {
        result[key] = key === 'SETTINGS' ? {} : [];
      }
    } else {
      result[key] = key === 'SETTINGS' ? {} : [];
    }
  });

  result.SETTINGS = Object.assign(getDefaultSettings_(), result.SETTINGS || {});
  return result;
}

function createSession_(idToko, isDemo) {
  const token = Utilities.getUuid();
  const payload = JSON.stringify({ idToko: idToko, isDemo: !!isDemo, createdAt: Date.now() });
  CacheService.getScriptCache().put('SESSION_' + token, payload, SESSION_TTL_SECONDS);
  return token;
}

function readSession_(sessionToken) {
  if (!sessionToken) throw new Error('Sesi login tidak ditemukan. Silakan login ulang.');
  const cached = CacheService.getScriptCache().get('SESSION_' + sessionToken);
  if (!cached) throw new Error('Sesi login berakhir. Silakan login ulang.');
  return JSON.parse(cached);
}

function verifikasiLoginServer(idToko) {
  setupDatabase();

  idToko = String(idToko || '').trim().toLowerCase();
  if (!idToko) return { success: false, message: 'ID Toko wajib diisi!' };

  const db = readDatabase_();
  const settings = db.SETTINGS || getDefaultSettings_();
  const idLogin = String(settings.idLogin || 'lunatic').trim().toLowerCase();
  const idDemo = String(settings.idDemo || 'demo').trim().toLowerCase();
  const isDemo = idToko === idDemo;

  if (idToko === idLogin || isDemo) {
    return {
      success: true,
      message: 'Login berhasil',
      settings: settings,
      isDemo: isDemo,
      sessionToken: createSession_(idToko, isDemo)
    };
  }

  return { success: false, message: 'ID Toko salah!' };
}

function loadInitialData(sessionToken) {
  readSession_(sessionToken);
  return readDatabase_();
}

function saveDataServer(sessionToken, key, jsonString) {
  const session = readSession_(sessionToken);
  if (session.isDemo) throw new Error('Mode Demo tidak boleh menyimpan data.');

  key = String(key || '').trim().toUpperCase();
  if (ALLOWED_KEYS.indexOf(key) === -1) {
    throw new Error('Key database tidak valid: ' + key);
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonString || (key === 'SETTINGS' ? '{}' : '[]'));
  } catch (err) {
    throw new Error('Format JSON tidak valid untuk ' + key);
  }

  if (key === 'SETTINGS') {
    const current = readDatabase_().SETTINGS || getDefaultSettings_();
    parsed = Object.assign({}, current, parsed);
  } else if (!Array.isArray(parsed)) {
    throw new Error('Data ' + key + ' harus berbentuk array.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSpreadsheet_();
    
    // Simpan ke sheet yang sesuai dengan key
    let sheet = ss.getSheetByName(key);
    if (!sheet) {
      sheet = ss.insertSheet(key);
      sheet.getRange('A1').setValue('Value JSON').setFontWeight('bold');
    }
    
    const safeJson = JSON.stringify(parsed);
    sheet.getRange('A2').setValue(safeJson);

    return { success: true, key: key, savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}
