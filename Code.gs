/**
 * Toonbank Cloud POS - Backend Google Apps Script
 * Database disimpan dalam sheet bernama "Database"
 */

const SHEET_NAME = 'Database';

function doGet(e) {
  // Setup kerangka database secara otomatis saat link web app diakses
  setupDatabase(); 
  
  // Memanggil file index.html (pastikan nama file di editor GAS adalah index.html huruf kecil)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Toonbank - Premium POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fungsi otomatis untuk membuat struktur database di Google Sheets jika belum ada
function setupDatabase() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Key', 'Value']);
    sheet.getRange("A1:B1").setFontWeight("bold");
    
    // Pengaturan Default (Termasuk ID Login, ID Demo, & Nama Toko)
    const defaultSettings = { idLogin: 'lunatic', idDemo: 'demo', namaToko: 'LUNACELL', namaApp: 'PREMIUM POS' };
    
    // Setup Baris Data Awal
    sheet.appendRow(['SETTINGS', JSON.stringify(defaultSettings)]);
    sheet.appendRow(['PRODUCTS', '[]']);
    sheet.appendRow(['WALLETS', '[]']);
    sheet.appendRow(['TRANSACTIONS', '[]']);
    sheet.appendRow(['HUTANG', '[]']);
    sheet.appendRow(['RIWAYAT_HUTANG', '[]']);
    sheet.appendRow(['MUTASI', '[]']);
  }
  return sheet;
}

// Fungsi mengecek login dari database
function verifikasiLoginServer(idToko) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || setupDatabase();
  let data = sheet.getDataRange().getValues();
  
  let settingsStr = '{}';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'SETTINGS') { 
      settingsStr = data[i][1]; 
      break; 
    }
  }
  
  let settings = JSON.parse(settingsStr);
  let isDemo = idToko === (settings.idDemo || 'demo');
  
  // Berikan akses jika ID Toko sama dengan ID Login Admin ATAU ID Demo
  if (idToko === settings.idLogin || isDemo) {
    return { success: true, message: 'Login berhasil', settings: settings, isDemo: isDemo };
  }
  return { success: false, message: 'ID Toko salah!' };
}

// Fungsi memuat semua data dari sheet saat berhasil login
function loadInitialData() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || setupDatabase();
  let data = sheet.getDataRange().getValues();
  let result = {};
  
  for (let i = 1; i < data.length; i++) {
    try {
      result[data[i][0]] = JSON.parse(data[i][1] || '[]');
    } catch(e) {
      result[data[i][0]] = (data[i][0] === 'SETTINGS') ? {} : [];
    }
  }
  return result;
}

// Fungsi utama untuk menyimpan/memperbarui data dari Frontend ke Sheet
function saveDataServer(key, jsonString) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || setupDatabase();
  let data = sheet.getDataRange().getValues();
  let found = false;
  
  // Cari baris berdasarkan Key dan update
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(jsonString); // Update value JSON
      found = true;
      break;
    }
  }
  
  // Jika key belum ada (fallback), buat baris baru
  if (!found) {
    sheet.appendRow([key, jsonString]);
  }
  
  return true;
}