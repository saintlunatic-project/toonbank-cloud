/**
 * BACKEND GOOGLE APPS SCRIPT (Code.gs)
 * ====================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('WarungKu POS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// SETUP DATABASE OTOMATIS
function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('WARUNGKU_DB_ID');
  let ss;

  try {
    if (ssId) ss = SpreadsheetApp.openById(ssId);
    else {
      ss = SpreadsheetApp.create('Database WarungKu POS');
      props.setProperty('WARUNGKU_DB_ID', ss.getId());
    }

    // 1. Sheet Produk
    let sheetProduk = ss.getSheetByName('Produk');
    if (!sheetProduk) {
      sheetProduk = ss.insertSheet('Produk');
      sheetProduk.appendRow(['ID', 'Nama', 'Harga Modal', 'Harga Jual', 'Stok', 'Kategori', 'Icon', 'IconBg']);
      sheetProduk.getRange("A1:H1").setFontWeight("bold").setBackground("#407a5b").setFontColor("white");
    }

    // 2. Sheet Transaksi (Ditambahkan kolom Pelanggan di akhir)
    let sheetTrx = ss.getSheetByName('Transaksi');
    if (!sheetTrx) {
      sheetTrx = ss.insertSheet('Transaksi');
      sheetTrx.appendRow(['ID Transaksi', 'Tanggal', 'Item', 'Total Bayar', 'Metode', 'Diskon', 'Pelanggan']);
      sheetTrx.getRange("A1:G1").setFontWeight("bold").setBackground("#407a5b").setFontColor("white");
    }

    // 3. Sheet Pengaturan (Baru untuk simpan ID Login dll)
    let sheetPengaturan = ss.getSheetByName('Pengaturan');
    if (!sheetPengaturan) {
      sheetPengaturan = ss.insertSheet('Pengaturan');
      sheetPengaturan.appendRow(['Key', 'Value']);
      sheetPengaturan.appendRow(['OwnerName', 'lunatic']);
      sheetPengaturan.appendRow(['StoreName', 'WarungKu']);
      sheetPengaturan.appendRow(['AdminID', 'lunatic']);
      sheetPengaturan.appendRow(['KasirID', 'kasir']);
      sheetPengaturan.getRange("A1:B1").setFontWeight("bold").setBackground("#407a5b").setFontColor("white");
    }
      
    const sheet1 = ss.getSheetByName('Sheet1');
    if (sheet1) ss.deleteSheet(sheet1);

    return "Setup Berhasil! URL Database: " + ss.getUrl();

  } catch (e) {
    return "Error Setup: " + e.message;
  }
}

// HELPER: Mengambil Database (Otomatis Setup jika belum ada)
function getDb() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('WARUNGKU_DB_ID');
  if (!ssId) {
    setupDatabase();
    ssId = props.getProperty('WARUNGKU_DB_ID');
  }
  const ss = SpreadsheetApp.openById(ssId);
  ensureDatabaseStructure(ss);
  return ss;
}

// Memastikan semua sheet wajib tetap ada, termasuk jika database lama/hasil copy belum lengkap.
function ensureDatabaseStructure(ss) {
  let sheetProduk = ss.getSheetByName('Produk');
  if (!sheetProduk) {
    sheetProduk = ss.insertSheet('Produk');
    sheetProduk.appendRow(['ID', 'Nama', 'Harga Modal', 'Harga Jual', 'Stok', 'Kategori', 'Icon', 'IconBg']);
    sheetProduk.getRange('A1:H1').setFontWeight('bold').setBackground('#407a5b').setFontColor('white');
  }

  let sheetTrx = ss.getSheetByName('Transaksi');
  if (!sheetTrx) {
    sheetTrx = ss.insertSheet('Transaksi');
    sheetTrx.appendRow(['ID Transaksi', 'Tanggal', 'Item', 'Total Bayar', 'Metode', 'Diskon', 'Pelanggan']);
    sheetTrx.getRange('A1:G1').setFontWeight('bold').setBackground('#407a5b').setFontColor('white');
  }

  let sheetPengaturan = ss.getSheetByName('Pengaturan');
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet('Pengaturan');
    sheetPengaturan.appendRow(['Key', 'Value']);
    sheetPengaturan.appendRow(['OwnerName', 'lunatic']);
    sheetPengaturan.appendRow(['StoreName', 'WarungKu']);
    sheetPengaturan.appendRow(['AdminID', 'lunatic']);
    sheetPengaturan.appendRow(['KasirID', 'kasir']);
    sheetPengaturan.getRange('A1:B1').setFontWeight('bold').setBackground('#407a5b').setFontColor('white');
  }

  const requiredSettings = { OwnerName: 'lunatic', StoreName: 'WarungKu', AdminID: 'lunatic', KasirID: 'kasir' };
  const data = sheetPengaturan.getDataRange().getValues();
  const existingKeys = data.slice(1).map(row => String(row[0] || ''));
  Object.keys(requiredSettings).forEach(key => {
    if (!existingKeys.includes(key)) sheetPengaturan.appendRow([key, requiredSettings[key]]);
  });
}

function getAppSettings() {
  const ss = getDb();
  const sheet = ss.getSheetByName('Pengaturan');
  const data = sheet.getDataRange().getValues();
  const settings = { ownerName: 'lunatic', storeName: 'WarungKu', adminId: 'lunatic', kasirId: 'kasir' };

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = String(data[i][1] || '').trim();
    if (key === 'OwnerName' && val) settings.ownerName = val;
    if (key === 'StoreName' && val) settings.storeName = val;
    if (key === 'AdminID' && val) settings.adminId = val;
    if (key === 'KasirID' && val) settings.kasirId = val;
  }

  return settings;
}

// --------------------------------------------------------
// FUNGSI CRUD (DIPANGGIL DARI FRONTEND)
// --------------------------------------------------------

function getInitialData() {
  const ss = getDb();
  
  // Ambil Produk
  const sheetProduk = ss.getSheetByName('Produk');
  const prodData = sheetProduk.getDataRange().getValues();
  const products = [];
  if (prodData.length > 1) {
    for (let i = 1; i < prodData.length; i++) {
      products.push({
        id: prodData[i][0], name: prodData[i][1], cost: prodData[i][2], price: prodData[i][3],
        stock: prodData[i][4], cat: prodData[i][5], icon: prodData[i][6], iconBg: prodData[i][7]
      });
    }
  }

  // Ambil Transaksi
  const sheetTrx = ss.getSheetByName('Transaksi');
  const trxData = sheetTrx.getDataRange().getValues();
  const transactions = [];
  if (trxData.length > 1) {
    for (let i = 1; i < trxData.length; i++) {
      transactions.push({
        id: trxData[i][0], date: trxData[i][1], items: JSON.parse(trxData[i][2] || "{}"),
        total: trxData[i][3], method: trxData[i][4], discount: trxData[i][5],
        customerName: trxData[i][6] || '' // Ditambahkan pembacaan pelanggan
      });
    }
  }

  // Ambil Pengaturan
  const settings = getAppSettings();
  
  return { products: products, transactions: transactions.reverse(), settings: settings };
}

function saveSettingsData(settings) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Pengaturan');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    let key = data[i][0];
    if (key === 'OwnerName') sheet.getRange(i + 1, 2).setValue(settings.ownerName);
    if (key === 'StoreName') sheet.getRange(i + 1, 2).setValue(settings.storeName);
    if (key === 'AdminID') sheet.getRange(i + 1, 2).setValue(settings.adminId);
    if (key === 'KasirID') sheet.getRange(i + 1, 2).setValue(settings.kasirId);
  }
  return true;
}

function saveProductData(p) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Produk');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == p.id) { rowIndex = i + 1; break; }
  }
  if(rowIndex > -1) sheet.getRange(rowIndex, 2, 1, 7).setValues([[p.name, p.cost, p.price, p.stock, p.cat, p.icon, p.iconBg]]);
  else sheet.appendRow([p.id, p.name, p.cost, p.price, p.stock, p.cat, p.icon, p.iconBg]);
  return true;
}

function deleteProductData(id) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Produk');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == id) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

function saveTransactionData(trx) {
  const ss = getDb();
  const sheetTrx = ss.getSheetByName('Transaksi');
  // Menambahkan trx.customerName ke kolom database
  sheetTrx.appendRow([trx.id, trx.date, JSON.stringify(trx.items), trx.total, trx.method, trx.discount, trx.customerName || '']);
  
  const sheetProduk = ss.getSheetByName('Produk');
  const prodData = sheetProduk.getDataRange().getValues();
  Object.keys(trx.items).forEach(itemId => {
    const qty = trx.items[itemId];
    for(let i = 1; i < prodData.length; i++) {
      if(prodData[i][0] == itemId) {
        let currentStock = prodData[i][4];
        sheetProduk.getRange(i + 1, 5).setValue(currentStock - qty);
        break;
      }
    }
  });
  return true;
}

function deleteTransactionData(id) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Transaksi');
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] == id) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

// Fungsi Hapus Transaksi Bulanan
function deleteMonthlyTransactionsData(monthStr) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Transaksi');
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i > 0; i--) {
    let dateVal = new Date(data[i][1]); 
    let m = ('0' + (dateVal.getMonth() + 1)).slice(-2);
    let y = dateVal.getFullYear();
    if (`${y}-${m}` === monthStr) {
      sheet.deleteRow(i + 1);
    }
  }
  return true;
}