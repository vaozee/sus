/**
 * ============================================================================
 *  KopSUS APP — BACKEND (Google Apps Script)
 * ============================================================================
 *  Spreadsheet sumber (DATABASE UTAMA — Inventory, Suppliers, Customers, PO,
 *  PD, SO, SD, Receipts, Payments, Dimensions, Users):
 *    https://docs.google.com/spreadsheets/d/1ExQlEKR5q4eMt0BwNanu7-KEPQdV-mD93seXNScveFg
 *
 *  Spreadsheet SIMPANAN VIEWER (data khusus anggota koperasi — Simpanan Pokok,
 *  Simpanan Wajib, Nilai SHU, Jabatan, NIK, Tgl Masuk Anggota — dipakai untuk
 *  dashboard role Viewer, dicocokkan dengan Customer ID dari database utama):
 *    https://docs.google.com/spreadsheets/d/1gKqxd5WTVBK90SkLIeix7rmkUxDg42Tti13kndYzKpU
 *
 *  Tab yang DIBACA di DATABASE UTAMA (harus sudah ada di spreadsheet dengan Named Range masing2):
 *    Sheet            Named Range
 *    ───────────────  ─────────────────────
 *    InventoryItems   RANGEINVENTORYITEMS
 *    Suppliers        RANGESUPPLIERS
 *    Customers        RANGECUSTOMERS
 *    PurchaseOrders   RANGEPO
 *    PurchaseDetails  RANGEPD
 *    SalesOrders      RANGESO
 *    SalesDetails     RANGESD
 *    Receipts         RANGERECEIPTS
 *    Payments         RANGEPAYMENTS
 *    Dimensions       RANGEDIMENSIONS
 *    Users            (dibaca by sheet name)
 *
 *  Tab yang DIBACA di SPREADSHEET SIMPANAN VIEWER:
 *    Sheet            Named Range
 *    ───────────────  ─────────────────────
 *    Customers        RANGECUSTOMERS (fallback: sheet bernama "Customers")
 *      → wajib berisi kolom "Customer ID" + kolom tambahan koperasi:
 *        Simpanan Pokok | Simpanan Wajib | Nilai SHU | Jabatan | NIK |
 *        Tgl Masuk Anggota. Baris dicocokkan ke database utama via Customer ID.
 *
 *  CARA DEPLOY:
 *  1. Buka https://script.google.com → New Project (atau buka dari Sheets:
 *     Extensions → Apps Script)
 *  2. Hapus isi default, paste SEMUA kode ini
 *  3. Jalankan `setupKopSUS` SEKALI dari editor untuk membuat tab Users
 *     dan user admin default jika belum ada.
 *     → Pertama kali Google akan minta otorisasi. Izinkan.
 *  4. Deploy → New deployment → tipe: Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Copy URL deployment (.../exec) → tempel ke konstanta BACKEND_URL
 *     di file index.html (Settings > Backend URL)
 *  6. Setiap edit kode: Manage deployments → Edit → New version
 * ============================================================================
 */

// ── KONFIGURASI ──────────────────────────────────────────────────────────────
// Database utama: Inventory, Suppliers, Customers, PO, PD, SO, SD, Receipts, Payments, Dimensions, Users
const SPREADSHEET_ID = '1ExQlEKR5q4eMt0BwNanu7-KEPQdV-mD93seXNScveFg';
// Spreadsheet Simpanan Viewer: kolom tambahan koperasi utk dashboard role Viewer
const VIEWER_SPREADSHEET_ID = '1gKqxd5WTVBK90SkLIeix7rmkUxDg42Tti13kndYzKpU';

// Sheet names
const SH_INVENTORY  = 'InventoryItems';
const SH_SUPPLIERS  = 'Suppliers';
const SH_CUSTOMERS  = 'Customers';
const SH_PO         = 'PurchaseOrders';
const SH_PD         = 'PurchaseDetails';
const SH_SO         = 'SalesOrders';
const SH_SD         = 'SalesDetails';
const SH_RECEIPTS   = 'Receipts';
const SH_PAYMENTS   = 'Payments';
const SH_DIMENSIONS = 'Dimensions';
const SH_USERS      = 'Users';

// Named range names (harus sudah dibuat di Sheets: Data → Named ranges)
const RNG_INVENTORY  = 'RANGEINVENTORYITEMS';
const RNG_SUPPLIERS  = 'RANGESUPPLIERS';
const RNG_CUSTOMERS  = 'RANGECUSTOMERS';
const RNG_PO         = 'RANGEPO';
const RNG_PD         = 'RANGEPD';
const RNG_SO         = 'RANGESO';
const RNG_SD         = 'RANGESD';
const RNG_RECEIPTS   = 'RANGERECEIPTS';
const RNG_PAYMENTS   = 'RANGEPAYMENTS';
const RNG_DIMENSIONS = 'RANGEDIMENSIONS';

// Kolom header untuk tiap sheet (urutan harus sama dengan sheet asli)
const COLS_INVENTORY = [
  'Item ID','Item Type','Item Category','Item Subcategory','Item Name',
  'QTY Purchased','QTY Sold','Remaining QTY','Reorder Level','Reorder Required'
];
const COLS_SUPPLIERS = [
  'Supplier ID','Supplier Name','Supplier Contact','Supplier Email',
  'State','City','Supplier Address',
  'Total Purchases','Total Payments','Balance Payable','Created By'
];
const COLS_CUSTOMERS = [
  'Customer ID','Customer Name','Customer Contact','Customer Email',
  'State','City','Customer Address',
  'Total Sales','Total Receipts','Balance Receivable','Created By',
  // Kolom khusus koperasi (idealnya diisi di spreadsheet Simpanan Viewer, sheet Customers)
  'Simpanan Pokok','Simpanan Wajib','Nilai SHU',
  'Jabatan','NIK','Tgl Masuk Anggota'
];
const COLS_PO = [
  'Date','PO ID','Supplier ID','Supplier Name','Bill Num',
  'State','City','Total Amount','Total Paid','PO Balance',
  'PMT Status','Shipping Status'
];
const COLS_PD = [
  'Date','PO ID','Supplier ID','Supplier Name','Bill Num','State','City',
  'Item ID','Item Name','Item Type','Item Category','Item Subcategory',
  'QTY Purchased','Unit Price','Total Purchase Price'
];
const COLS_SO = [
  'SO Date','SO ID','Customer ID','Customer Name','Invoice Num',
  'State','City','Total SO Amount','Total Received','SO Balance','Receipt Status'
];
const COLS_SD = [
  'SO Date','SO ID','Detail ID','Customer ID','Customer Name','State','City',
  'Invoice Num','Item Name','Item ID','Item Type','Item Category','Item Subcategory',
  'QTY Sold','Unit Price','Price Incl Tax','Shipping Fees','Total Sales Price'
];
const COLS_RECEIPTS = [
  'Trx Date','Trx ID','Customer ID','Customer Name','State','City',
  'SO ID','Invoice Num','PMT Mode','Amount Received'
];
const COLS_PAYMENTS = [
  'Trx Date','Trx ID','Supplier ID','Supplier Name','State','City',
  'PO ID','Bill Num','PMT Mode','Amount Paid'
];
const COLS_DIMENSIONS = ['Column Name','Value'];
const COLS_USERS = [
  'User ID','Name','Email','Role','Status','Date Added',
  'Username','Password Hash','Linked Customer ID'
];

// Auth
const AUTH_SALT = 'KopSUS-2026-salt';
const AUTH_SESSION_SEC = 8 * 60 * 60; // 8 jam

// ── ENTRY POINTS ─────────────────────────────────────────────────────────────
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }
function doOptions(e) { return ContentService.createTextOutput(''); }

function handleRequest(e) {
  try {
    const body    = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    // PENTING: action bisa datang dari query string (GET) ATAU dari body JSON (POST,
    // dipakai untuk action seperti 'login' yang di-POST tanpa query string di URL).
    const action  = (e.parameter.action || body.action || '').toString();
    const payload = Object.assign({}, e.parameter, body);

    let result;
    switch (action) {
      // ── AUTH ──
      case 'login':              result = login(payload);              break;
      case 'logout':             result = logout(payload);             break;
      case 'validateSession':    result = validateSession(payload);    break;

      // ── DASHBOARD ──
      case 'getDashboard':         result = getDashboard(payload);         break;
      case 'getViewerDashboard':   result = getViewerDashboard(payload);   break;

      // ── INVENTORY ──
      case 'getInventory':       result = getInventory(payload);       break;
      case 'addInventory':       result = addInventory(payload);       break;
      case 'updateInventory':    result = updateInventory(payload);    break;
      case 'deleteInventory':    result = deleteInventory(payload);    break;

      // ── DIMENSIONS ──
      case 'getDimensions':      result = getDimensions(payload);      break;
      case 'addDimension':       result = addDimension(payload);       break;
      case 'deleteDimension':    result = deleteDimension(payload);    break;

      // ── SUPPLIERS ──
      case 'getSuppliers':       result = getSuppliers(payload);       break;
      case 'addSupplier':        result = addSupplier(payload);        break;
      case 'updateSupplier':     result = updateSupplier(payload);     break;
      case 'deleteSupplier':     result = deleteSupplier(payload);     break;

      // ── CUSTOMERS ──
      case 'getCustomers':       result = getCustomers(payload);       break;
      case 'addCustomer':        result = addCustomer(payload);        break;
      case 'updateCustomer':     result = updateCustomer(payload);     break;
      case 'deleteCustomer':     result = deleteCustomer(payload);     break;

      // ── PURCHASE ORDERS ──
      case 'getPurchaseOrders':  result = getPurchaseOrders(payload);  break;
      case 'addPurchaseOrder':   result = addPurchaseOrder(payload);   break;
      case 'updatePurchaseOrder':result = updatePurchaseOrder(payload);break;
      case 'deletePurchaseOrder':result = deletePurchaseOrder(payload);break;

      // ── PURCHASE DETAILS ──
      case 'getPurchaseDetails': result = getPurchaseDetails(payload); break;
      case 'addPurchaseDetail':  result = addPurchaseDetail(payload);  break;
      case 'deletePurchaseDetail':result= deletePurchaseDetail(payload);break;

      // ── SALES ORDERS ──
      case 'getSalesOrders':     result = getSalesOrders(payload);     break;
      case 'addSalesOrder':      result = addSalesOrder(payload);      break;
      case 'updateSalesOrder':   result = updateSalesOrder(payload);   break;
      case 'deleteSalesOrder':   result = deleteSalesOrder(payload);   break;

      // ── SALES DETAILS ──
      case 'getSalesDetails':    result = getSalesDetails(payload);    break;
      case 'addSalesDetail':     result = addSalesDetail(payload);     break;
      case 'deleteSalesDetail':  result = deleteSalesDetail(payload);  break;

      // ── RECEIPTS ──
      case 'getReceipts':        result = getReceipts(payload);        break;
      case 'addReceipt':         result = addReceipt(payload);         break;
      case 'updateReceipt':      result = updateReceipt(payload);      break;
      case 'deleteReceipt':      result = deleteReceipt(payload);      break;

      // ── PAYMENTS ──
      case 'getPayments':        result = getPayments(payload);        break;
      case 'addPayment':         result = addPayment(payload);         break;
      case 'updatePayment':      result = updatePayment(payload);      break;
      case 'deletePayment':      result = deletePayment(payload);      break;

      // ── USERS (Settings) ──
      case 'getUsers':           result = getUsers(payload);           break;
      case 'addUser':            result = addUser(payload);            break;
      case 'updateUser':         result = updateUser(payload);         break;
      case 'deleteUser':         result = deleteUser(payload);         break;
      case 'setUserPassword':    result = setUserPassword(payload);    break;

      // ── UTIL ──
      case 'ping':               result = { ok: true, time: new Date().toISOString() }; break;
      case 'setupNamedRanges':   result = setupNamedRanges();          break;

      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message, stack: err.stack });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SETUP ────────────────────────────────────────────────────────────────────
/**
 * Jalankan SEKALI dari Apps Script Editor.
 * Membuat tab Users (jika belum ada) + user admin default.
 * Juga membuat/memperbarui Named Ranges untuk semua tab.
 */
function setupKopSUS() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheet(ss, SH_USERS, COLS_USERS);

  // Buat user admin default jika Users masih kosong
  const usersSheet = ss.getSheetByName(SH_USERS);
  if (usersSheet.getLastRow() <= 1) {
    const userId = 'USR-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const hash   = hashPassword('admin123');
    usersSheet.appendRow([
      userId, 'Administrator', 'admin@kopsus.id',
      'Admin', 'Active', new Date().toISOString().substring(0, 10),
      'admin', hash, ''
    ]);
    Logger.log('✅ User admin dibuat. Username: admin / Password: admin123 — SEGERA GANTI!');
  } else {
    Logger.log('ℹ️  Tab Users sudah ada, tidak dibuat ulang.');
  }

  // Buat named ranges untuk semua tab
  setupNamedRanges();
  Logger.log('✅ Setup KopSUS selesai.');
  SpreadsheetApp.getUi().alert(
    '✅ Setup KopSUS selesai!\n\n' +
    'User admin:\n  Username: admin\n  Password: admin123\n\n' +
    'Named ranges berhasil dibuat untuk semua tab.\n' +
    'Silakan deploy sebagai Web App lalu copy URL-nya ke aplikasi frontend.'
  );
}

/**
 * Buat/perbarui Named Ranges untuk semua tab utama.
 * Named Range mencakup seluruh isi sheet (header + data).
 * Jalankan ulang setiap kali ada baris baru agar range tidak ketinggalan.
 * (Alternatif lebih baik: gunakan range terbuka mis. A:Z di Named Range via UI Sheets)
 */
function setupNamedRanges() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pairs = [
    [SH_INVENTORY,  RNG_INVENTORY],
    [SH_SUPPLIERS,  RNG_SUPPLIERS],
    [SH_CUSTOMERS,  RNG_CUSTOMERS],
    [SH_PO,         RNG_PO],
    [SH_PD,         RNG_PD],
    [SH_SO,         RNG_SO],
    [SH_SD,         RNG_SD],
    [SH_RECEIPTS,   RNG_RECEIPTS],
    [SH_PAYMENTS,   RNG_PAYMENTS],
    [SH_DIMENSIONS, RNG_DIMENSIONS],
  ];

  const existingNR = {};
  ss.getNamedRanges().forEach(nr => { existingNR[nr.getName()] = nr; });

  pairs.forEach(([sheetName, rangeName]) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { Logger.log('⚠️  Sheet "' + sheetName + '" tidak ditemukan, named range dilewati.'); return; }
    const lastRow = Math.max(sh.getLastRow(), 2);
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const rng = sh.getRange(1, 1, lastRow, lastCol);
    if (existingNR[rangeName]) {
      existingNR[rangeName].setRange(rng);
      Logger.log('🔄 Named range "' + rangeName + '" diperbarui.');
    } else {
      ss.setNamedRange(rangeName, rng);
      Logger.log('✅ Named range "' + rangeName + '" dibuat.');
    }
  });

  return { ok: true, message: 'Named ranges berhasil dibuat/diperbarui.' };
}

function ensureSheet(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(cols);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('✅ Tab "' + name + '" dibuat dengan header.');
  }
  return sh;
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Spreadsheet terpisah berisi data Simpanan (khusus dashboard role Viewer)
function getViewerSS() {
  return SpreadsheetApp.openById(VIEWER_SPREADSHEET_ID);
}

/**
 * Cari 1 baris di sheet "Customers" pada spreadsheet Simpanan Viewer,
 * dicocokkan dengan Customer ID. Coba named range RANGECUSTOMERS dulu,
 * fallback ke sheet biasa bernama "Customers".
 */
// Normalisasi ID: trim + uppercase, supaya "cust001", "CUST001 ", "Cust001" dianggap sama.
function normId(v) {
  return (v === undefined || v === null) ? '' : v.toString().trim().toUpperCase();
}

function getViewerSimpananRow(customerId) {
  let vss;
  try { vss = getViewerSS(); } catch (e) { return {}; }

  // Sumber utama: tab "Sheet1" pada spreadsheet Simpanan Viewer.
  // Header aktual: No, Code Pelanggan, NIK, Nama, Jabatan, Hit SHU,
  // Tanggal Masuk Anggota, Simpanan Pokok, Simpanan Wajib, Belanja Bulan ini, Nilai SHU
  let sh = vss.getSheetByName('Sheet1');
  if (!sh) {
    // Fallback ke skema lama (named range / sheet "Customers"), untuk kompatibilitas.
    try {
      const rng = vss.getRangeByName(RNG_CUSTOMERS);
      if (rng) {
        const vals = rng.getValues();
        return matchSimpananRow_(vals, customerId, ['Customer ID', 'Code Pelanggan']);
      }
    } catch (e) { /* ignore */ }
    const plain = vss.getSheetByName(SH_CUSTOMERS);
    if (!plain) return {};
    return matchSimpananRow_(plain.getDataRange().getValues(), customerId, ['Customer ID', 'Code Pelanggan']);
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1) return {};
  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  return matchSimpananRow_(vals, customerId, ['Code Pelanggan', 'Customer ID']);
}

// Helper: cocokkan baris berdasarkan salah satu dari beberapa kemungkinan nama kolom ID.
function matchSimpananRow_(vals, customerId, idColCandidates) {
  if (!vals || vals.length < 1) return {};
  const headers = vals[0].map(h => String(h).trim());
  let iId = -1;
  for (const cand of idColCandidates) {
    iId = headers.indexOf(cand);
    if (iId !== -1) break;
  }
  if (iId === -1) return {};
  const target = normId(customerId);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (normId(row[iId]) !== target) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    return obj;
  }
  return {};
}

/**
 * Ambil semua baris item-belanja milik seorang anggota dari tab "Rincian"
 * pada spreadsheet Simpanan Viewer, lalu group per RESI (1 RESI = 1 transaksi/nota,
 * bisa berisi banyak baris NAMA BARANG).
 * Header aktual tab Rincian: CODE, Code Pelanggan, NO., RESI, KODE PELANGGAN, COB,
 * NAMA PELANGGAN, NAMA BARANG, ALAMAT PELANGGAN, SATUAN, QTY, HARGA, TOTAL, JUMLAH
 */
function getViewerRincianRows(customerId) {
  let vss;
  try { vss = getViewerSS(); } catch (e) { return []; }

  const sh = vss.getSheetByName('Rincian');
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = vals[0].map(h => String(h).trim());
  const iCode = headers.indexOf('Code Pelanggan');
  if (iCode === -1) return [];
  const target = normId(customerId);

  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (normId(row[iCode]) !== target) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    rows.push(obj);
  }
  return rows;
}

/**
 * Group baris-baris Rincian (per item) menjadi transaksi (per RESI).
 * Return array { resi, namaPelanggan, alamatPelanggan, amountReceived, items:[...] }
 * diurutkan sesuai urutan kemunculan di sheet (biasanya = urutan tanggal, terbaru di bawah).
 */
function groupRincianByResi_(rows) {
  const map = {};
  const order = [];
  rows.forEach(r => {
    const resi = (r['RESI'] || '').toString().trim() || '(tanpa resi)';
    if (!map[resi]) {
      map[resi] = {
        resi,
        namaPelanggan: r['NAMA PELANGGAN'] || '',
        alamatPelanggan: r['ALAMAT PELANGGAN'] || '',
        amountReceived: 0,
        items: []
      };
      order.push(resi);
    }
    const total = parseNum(r['TOTAL']);
    map[resi].amountReceived += total;
    map[resi].items.push({
      itemName: r['NAMA BARANG'] || '',
      satuan:   r['SATUAN'] || '',
      qty:      parseNum(r['QTY']),
      harga:    parseNum(r['HARGA']),
      total
    });
  });
  return order.map(resi => map[resi]).reverse(); // terbaru duluan
}

/**
 * Baca seluruh isi named range, return array of object {header: value}.
 * Tiap row juga mendapat properti _row (nomor baris di sheet, 1-based) untuk keperluan update/delete.
 */
function rangeToObjects(rangeName) {
  const ss = getSS();
  const rng = ss.getRangeByName(rangeName);
  if (!rng) throw new Error('Named range "' + rangeName + '" tidak ditemukan. Jalankan setupKopSUS() atau setupNamedRanges() terlebih dahulu.');
  const values = rng.getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const result  = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(v => v === '' || v === null || v === undefined)) continue;
    const obj = { _row: i + 1 }; // +1 karena header di baris 1, data mulai baris 2
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    result.push(obj);
  }
  return result;
}

/**
 * Tambah baris baru ke sheet (append di bawah data terakhir).
 */
function appendToSheet(sheetName, cols, data) {
  const ss   = getSS();
  const sh   = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');
  const row = cols.map(c => data[c] !== undefined ? data[c] : '');
  sh.appendRow(row);
  // Perbarui named range agar mencakup baris baru
  refreshNamedRange(ss, sheetName);
}

/**
 * Update satu baris berdasarkan _row (nomor baris di sheet, 1-based).
 * Baris 1 = header, baris 2 = data pertama.
 */
function updateSheetRow(sheetName, cols, rowNum, data) {
  const ss = getSS();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');
  // rowNum yang dikirim = _row dari rangeToObjects, yaitu index dalam range (1-based dari header)
  // di sheet asli: baris data = rowNum + 1 (karena baris 1 = header di sheet)
  const sheetRow = rowNum + 1;
  const row = cols.map(c => data[c] !== undefined ? data[c] : '');
  sh.getRange(sheetRow, 1, 1, cols.length).setValues([row]);
}

/**
 * Hapus satu baris berdasarkan _row.
 */
function deleteSheetRow(sheetName, rowNum) {
  const ss = getSS();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');
  const sheetRow = rowNum + 1;
  sh.deleteRow(sheetRow);
  refreshNamedRange(ss, sheetName);
}

/**
 * Perbarui named range setelah append/delete agar baris baru tercover.
 */
function refreshNamedRange(ss, sheetName) {
  const nameMap = {
    [SH_INVENTORY]: RNG_INVENTORY, [SH_SUPPLIERS]: RNG_SUPPLIERS,
    [SH_CUSTOMERS]: RNG_CUSTOMERS, [SH_PO]: RNG_PO, [SH_PD]: RNG_PD,
    [SH_SO]: RNG_SO, [SH_SD]: RNG_SD, [SH_RECEIPTS]: RNG_RECEIPTS,
    [SH_PAYMENTS]: RNG_PAYMENTS, [SH_DIMENSIONS]: RNG_DIMENSIONS
  };
  const rangeName = nameMap[sheetName];
  if (!rangeName) return;
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastRow = Math.max(sh.getLastRow(), 2);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const nrs = ss.getNamedRanges();
  const nr  = nrs.find(n => n.getName() === rangeName);
  if (nr) nr.setRange(sh.getRange(1, 1, lastRow, lastCol));
}

function parseNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function hashPassword(plain) {
  const input  = AUTH_SALT + ':' + plain;
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function generateId(prefix) {
  return prefix + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}

// ── AUTH ─────────────────────────────────────────────────────────────────────
function login(p) {
  const username  = (p.username || '').toString().trim().toLowerCase();
  const password  = (p.password || '').toString();
  if (!username || !password) return { ok: false, error: 'Username dan password wajib diisi.' };

  const ss    = getSS();
  const sh    = ss.getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'Belum ada user terdaftar.' };

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();

  const iUser = headers.indexOf('Username');
  const iHash = headers.indexOf('Password Hash');
  const iName = headers.indexOf('Name');
  const iRole = headers.indexOf('Role');
  const iStatus = headers.indexOf('Status');
  const iId   = headers.indexOf('User ID');
  const iLinked = headers.indexOf('Linked Customer ID');

  const hash = hashPassword(password);

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowUser = (row[iUser] || '').toString().trim().toLowerCase();
    if (rowUser !== username) continue;

    if ((row[iStatus] || '').toString() !== 'Active') {
      return { ok: false, error: 'Akun tidak aktif. Hubungi administrator.' };
    }
    if ((row[iHash] || '').toString().trim() !== hash) {
      return { ok: false, error: 'Username atau password salah.' };
    }

    const userData = {
      id: row[iId] || '', name: row[iName] || '', email: headers.indexOf('Email') > -1 ? row[headers.indexOf('Email')] : '',
      role: row[iRole] || '', username: row[iUser] || '',
      linkedCustomerId: iLinked > -1 ? (row[iLinked] || '').toString().trim() : ''
    };

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('ks_session_' + token, JSON.stringify(userData), AUTH_SESSION_SEC);

    // Update Last Login
    const iLastLogin = headers.indexOf('Date Added');
    // Tidak update Date Added, cukup log saja
    Logger.log('Login: ' + username + ' at ' + new Date().toISOString());

    return { ok: true, token, user: userData };
  }
  return { ok: false, error: 'Username atau password salah.' };
}

function logout(p) {
  if (p.token) CacheService.getScriptCache().remove('ks_session_' + p.token);
  return { ok: true };
}

function validateSession(p) {
  if (!p.token) return { ok: false, error: 'Token diperlukan.' };
  const cached = CacheService.getScriptCache().get('ks_session_' + p.token);
  if (!cached) return { ok: false, error: 'Session expired atau tidak valid.' };
  try {
    const user = JSON.parse(cached);
    // Perpanjang session
    CacheService.getScriptCache().put('ks_session_' + p.token, cached, AUTH_SESSION_SEC);
    return { ok: true, user };
  } catch (e) {
    return { ok: false, error: 'Session data tidak valid.' };
  }
}

// ── DASHBOARD ────────────────────────────────────────────────────────────────
function getDashboard(p) {
  const sales     = rangeToObjects(RNG_SD);
  const purchases = rangeToObjects(RNG_PD);
  const customers = rangeToObjects(RNG_CUSTOMERS);
  const suppliers = rangeToObjects(RNG_SUPPLIERS);

  const totalSales     = sales.reduce((s, r) => s + parseNum(r['Total Sales Price']), 0);
  const totalPurchases = purchases.reduce((s, r) => s + parseNum(r['Total Purchase Price']), 0);
  const netProfit      = totalSales - totalPurchases;
  const totalReceivable = customers.reduce((s, r) => s + parseNum(r['Balance Receivable']), 0);
  const totalPayable    = suppliers.reduce((s, r) => s + parseNum(r['Balance Payable']), 0);

  // Top location & item
  const cityMap  = groupSum(sales, 'City', 'Total Sales Price');
  const topLocation = topKey(cityMap);
  const itemMap  = groupSum(sales, 'Item Type', 'Total Sales Price');
  const topItem  = topKey(itemMap);

  // Chart 1: Sales Trend (by month YYYY-MM)
  const trendMap = {};
  sales.forEach(r => {
    const d = tryDate(r['SO Date']);
    const key = d ? fmtYM(d) : 'unknown';
    trendMap[key] = (trendMap[key] || 0) + parseNum(r['Total Sales Price']);
  });
  const trendDates  = Object.keys(trendMap).filter(k => k !== 'unknown').sort();
  const trendValues = trendDates.map(k => trendMap[k]);

  // Chart 2: Sales By State
  const stateMap = groupSum(sales, 'State', 'Total Sales Price');

  // Chart 3: Sales By Category (Item Type) - as percentage
  const catMap   = groupSum(sales, 'Item Type', 'Total Sales Price');
  const catTotal = Object.values(catMap).reduce((a, b) => a + b, 0) || 1;
  const salesByCategory = {
    labels: Object.keys(catMap),
    values: Object.values(catMap).map(v => Math.round(v / catTotal * 10000) / 100)
  };

  // Chart 4: Top 10 Customers
  const custSalesMap = groupSum(sales, 'Customer Name', 'Total Sales Price');
  const top10 = Object.entries(custSalesMap).sort((a, b) => b[1] - a[1]).slice(0, 10);

  // Chart 5: Purchase By Location (State) - donut
  const purStateMap = groupSum(purchases, 'State', 'Total Purchase Price');

  // Chart 6: Purchase By Category stacked by year
  const purCatYear = {};
  purchases.forEach(r => {
    const d = tryDate(r['Date']);
    const y = d ? d.getFullYear() : 'Unknown';
    const c = r['Item Type'] || 'Unknown';
    purCatYear[y] = purCatYear[y] || {};
    purCatYear[y][c] = (purCatYear[y][c] || 0) + parseNum(r['Total Purchase Price']);
  });
  const purYears   = Object.keys(purCatYear).sort();
  const purItems   = [...new Set(purchases.map(r => r['Item Type'] || 'Unknown'))];
  const purSeries  = purItems.map(item => ({
    name: item, data: purYears.map(y => purCatYear[y][item] || 0)
  }));

  // Chart 7: Sales By City (Treemap)
  const treemap = Object.entries(cityMap).sort((a, b) => b[1] - a[1]).map(([x, y]) => ({ x, y }));

  return {
    ok: true,
    kpi: { totalSales, totalPurchases, netProfit, totalReceivable, totalPayable, topLocation, topItem },
    charts: {
      salesTrend: { dates: trendDates, values: trendValues },
      salesByLocation: { labels: Object.keys(stateMap), values: Object.values(stateMap) },
      salesByCategory,
      topCustomers: { labels: top10.map(a => a[0]), values: top10.map(a => a[1]) },
      purchaseByLocation: { labels: Object.keys(purStateMap), values: Object.values(purStateMap) },
      purchaseByCategory: { years: purYears, series: purSeries },
      salesByCity: treemap
    }
  };
}

function groupSum(rows, keyField, valField) {
  const map = {};
  rows.forEach(r => {
    const k = (r[keyField] || 'Unknown').toString();
    map[k] = (map[k] || 0) + parseNum(r[valField]);
  });
  return map;
}

function topKey(map) {
  const e = Object.entries(map).sort((a, b) => b[1] - a[1])[0];
  return e ? e[0] : '';
}

function tryDate(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function fmtYM(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

// ── VIEWER DASHBOARD ─────────────────────────────────────────────────────────
/**
 * Dashboard khusus untuk role Viewer (anggota koperasi yang ter-link ke Customer ID).
 * Mengembalikan:
 *   - Data profil anggota (dari sheet Customers: Simpanan Pokok, Simpanan Wajib,
 *     Nilai SHU, Jabatan, NIK, Tgl Masuk Anggota, dll.)
 *   - KPI belanja: Belanja Bulan Ini, Total Belanja
 *   - Chart: Produk Tersering Dibeli (Top 10), Total Belanja Terbesar per Produk (Donut),
 *     Tren Belanja per Transaksi (Area), Rincian Belanja (Tabel Receipts + SD)
 *
 * Kolom extra di sheet Customers yang dibaca (jika ada):
 *   Simpanan Pokok | Simpanan Wajib | Nilai SHU | Jabatan | NIK | Tgl Masuk Anggota
 * Jika kolom tidak ada di sheet, value akan '' / 0.
 */
function getViewerDashboard(p) {
  const token = p.token || '';
  if (!token) return { ok: false, error: 'Token diperlukan.' };

  // Validasi session & ambil linkedCustomerId
  const cached = CacheService.getScriptCache().get('ks_session_' + token);
  if (!cached) return { ok: false, error: 'Session expired. Silakan login ulang.' };
  const user = JSON.parse(cached);
  const customerId = (user.linkedCustomerId || '').toString().trim();
  if (!customerId) return { ok: false, error: 'Akun ini tidak ter-link ke data anggota.' };

  const ss = getSS();

  // ── 1. Profil anggota dari sheet Customers ──
  let member = {};
  try {
    const custRange = ss.getRangeByName(RNG_CUSTOMERS);
    if (custRange) {
      const vals = custRange.getValues();
      const headers = vals[0].map(h => String(h).trim());
      const iId = headers.indexOf('Customer ID');
      const targetId = normId(customerId);
      for (let i = 1; i < vals.length; i++) {
        const row = vals[i];
        if (normId(row[iId]) !== targetId) continue;
        headers.forEach((h, j) => { member[h] = row[j] !== undefined ? row[j] : ''; });
        break;
      }
    }
  } catch (e) { /* tolerate */ }

  // ── 1b. Data Simpanan (Simpanan Pokok/Wajib/SHU dll) dari spreadsheet Viewer terpisah ──
  let simpanan = {};
  try { simpanan = getViewerSimpananRow(customerId); } catch (e) { /* tolerate */ }

  // Jika Customer ID tidak ditemukan sama sekali di sheet Customers utama MAUPUN
  // di spreadsheet Simpanan Viewer, hentikan lebih awal dengan pesan yang jelas
  // (bukannya diam-diam menampilkan dashboard kosong berisi Rp 0 / "–").
  const memberEmpty   = Object.keys(member).length === 0;
  const simpananEmpty = Object.keys(simpanan).length === 0;
  if (memberEmpty && simpananEmpty) {
    return {
      ok: false,
      error: 'Customer ID "' + customerId + '" (dari Linked Customer ID akun ini) tidak ditemukan ' +
        'di sheet Customers database utama maupun di spreadsheet Simpanan Viewer. ' +
        'Mohon periksa kembali kolom "Linked Customer ID" pada sheet Users, dan pastikan nilainya ' +
        'sama persis dengan "Customer ID" pada sheet Customers.'
    };
  }

  // Helper: ambil nilai kolom khusus koperasi, toleran jika tidak ada.
  // Prioritas: spreadsheet Simpanan Viewer, fallback ke sheet Customers database utama.
  const mVal = (k) => (simpanan[k] !== undefined && simpanan[k] !== '') ? simpanan[k] : (member[k] !== undefined ? member[k] : '');
  const mNum = (k) => parseNum(mVal(k));

  // ── 2. Riwayat belanja anggota ini dari tab "Rincian" (spreadsheet Simpanan Viewer) ──
  let rincianRaw = [];
  try { rincianRaw = getViewerRincianRows(customerId); } catch (e) { /* tolerate */ }
  let rincian = groupRincianByResi_(rincianRaw);

  // Fallback: kalau tidak ada apa-apa di tab Rincian, coba skema lama (SalesDetails+Receipts
  // di spreadsheet utama), untuk anggota yang datanya masih berada di sistem lama.
  if (!rincianRaw.length) {
    let sd = [], receipts = [];
    try {
      const sdRange = ss.getRangeByName(RNG_SD);
      if (sdRange) {
        const vals = sdRange.getValues();
        if (vals.length > 1) {
          const headers = vals[0].map(h => String(h).trim());
          const iCid = headers.indexOf('Customer ID');
          const targetIdSd = normId(customerId);
          for (let i = 1; i < vals.length; i++) {
            const row = vals[i];
            if (row.every(v => v === '' || v === null)) continue;
            if (normId(row[iCid]) !== targetIdSd) continue;
            const obj = {};
            headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
            sd.push(obj);
          }
        }
      }
    } catch (e) { /* tolerate */ }
    try {
      const rcRange = ss.getRangeByName(RNG_RECEIPTS);
      if (rcRange) {
        const vals = rcRange.getValues();
        if (vals.length > 1) {
          const headers = vals[0].map(h => String(h).trim());
          const iCid = headers.indexOf('Customer ID');
          const targetIdRc = normId(customerId);
          for (let i = 1; i < vals.length; i++) {
            const row = vals[i];
            if (row.every(v => v === '' || v === null)) continue;
            if (normId(row[iCid]) !== targetIdRc) continue;
            const obj = {};
            headers.forEach((h, j) => {
              let v = row[j];
              if (String(h).endsWith('Date') && v instanceof Date) {
                v = Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
              }
              obj[h] = v !== undefined ? v : '';
            });
            receipts.push(obj);
          }
        }
      }
    } catch (e) { /* tolerate */ }

    if (receipts.length || sd.length) {
      rincian = receipts.map(rc => {
        const soId = (rc['SO ID'] || '').toString().trim();
        const items = sd.filter(d => (d['SO ID'] || '').toString().trim() === soId);
        return {
          resi: rc['Trx ID'] || soId,
          trxDate: rc['Trx Date'],
          pmtMode: rc['PMT Mode'],
          amountReceived: parseNum(rc['Amount Received']),
          items: items.map(d => ({
            itemName: d['Item Name'] || '',
            satuan: d['Item Subcategory'] || d['Item Category'] || '',
            qty: parseNum(d['QTY Sold']),
            harga: parseNum(d['Unit Price']),
            total: parseNum(d['Total Sales Price'])
          }))
        };
      }).sort((a, b) => (tryDate(b.trxDate) || 0) - (tryDate(a.trxDate) || 0));
      rincianRaw = sd; // dipakai sekadar penanda "ada data" untuk kpi/chart fallback di bawah
    }
  }

  // ── 3. KPI ──
  // "Belanja Bulan Ini" sudah computed manual di kolom Sheet1, pakai langsung (lebih akurat
  // daripada dihitung ulang dari tanggal transaksi, karena tab Rincian tidak punya kolom tanggal).
  const belanjaMonth = mNum('Belanja Bulan ini');
  const totalBelanja = rincian.reduce((s, rc) => s + (rc.amountReceived || 0), 0);
  const totalItem = rincian.reduce((s, rc) => s + rc.items.reduce((si, it) => si + (it.qty || 0), 0), 0);

  // ── 4. Chart 1: Produk Tersering Dibeli (Top 10 by QTY) ──
  const itemQtyMap = {};
  rincianRaw.forEach(r => {
    const item = (r['NAMA BARANG'] || r['Item Name'] || 'Unknown').toString().trim();
    itemQtyMap[item] = (itemQtyMap[item] || 0) + parseNum(r['QTY'] !== undefined ? r['QTY'] : r['QTY Sold']);
  });
  const topItems = Object.entries(itemQtyMap)
    .sort((a, b) => b[1] - a[1]).slice(0, 10);
  const chartTopItems = {
    labels: topItems.map(a => a[0]),
    values: topItems.map(a => a[1])
  };

  // ── 5. Chart 2: Total Belanja Terbesar per Produk (Donut by TOTAL) ──
  const itemAmtMap = {};
  rincianRaw.forEach(r => {
    const item = (r['NAMA BARANG'] || r['Item Name'] || 'Unknown').toString().trim();
    itemAmtMap[item] = (itemAmtMap[item] || 0) + parseNum(r['TOTAL'] !== undefined ? r['TOTAL'] : r['Total Sales Price']);
  });
  const chartDonut = {
    labels: Object.keys(itemAmtMap),
    values: Object.values(itemAmtMap)
  };

  // ── 6. Chart 3: Tren Belanja per Transaksi (per RESI, urut sesuai kemunculan di sheet) ──
  const trendArr = rincian.slice().reverse(); // urut kronologis (lama → baru) untuk chart tren
  const chartTrend = {
    labels: trendArr.map(rc => rc.resi),
    values: trendArr.map(rc => rc.amountReceived)
  };

  return {
    ok: true,
    member: {
      customerId,
      name:          mVal('Nama') || mVal('Customer Name'),
      contact:       mVal('Customer Contact'),
      email:         mVal('Customer Email'),
      state:         mVal('State'),
      city:          mVal('City'),
      address:       mVal('Customer Address'),
      totalSales:    mNum('Total Sales'),
      totalReceipts: mNum('Total Receipts'),
      balance:       mNum('Balance Receivable'),
      // Kolom khusus koperasi (dari tab "Sheet1" spreadsheet Simpanan Viewer)
      simpananPokok:  mVal('Simpanan Pokok')  || mNum('Simpanan Pokok'),
      simpananWajib:  mVal('Simpanan Wajib')  || mNum('Simpanan Wajib'),
      nilaiSHU:       mVal('Nilai SHU')       || mNum('Nilai SHU'),
      jabatan:        mVal('Jabatan'),
      nik:            mVal('NIK'),
      tglMasuk:       mVal('Tanggal Masuk Anggota') || mVal('Tgl Masuk Anggota') || mVal('Tgl Masuk') || mVal('Join Date'),
    },
    kpi: {
      belanjaMonth,
      totalBelanja,
      totalTransaksi: rincian.length,
      totalItem
    },
    charts: { topItems: chartTopItems, donut: chartDonut, trend: chartTrend },
    rincian
  };
}

function getInventory(p) {
  let rows = rangeToObjects(RNG_INVENTORY);
  if (p.itemType)     rows = rows.filter(r => r['Item Type'] === p.itemType);
  if (p.itemCategory) rows = rows.filter(r => r['Item Category'] === p.itemCategory);
  const NUM = ['QTY Purchased','QTY Sold','Remaining QTY','Reorder Level'];
  rows = rows.map(r => {
    const o = Object.assign({}, r);
    NUM.forEach(k => { o[k] = parseNum(o[k]); });
    o['Reorder Required'] = String(o['Reorder Required']).toLowerCase() === 'true' ||
      o['Reorder Required'] === true || o['Reorder Required'] === 'Ya';
    return o;
  });
  return { ok: true, count: rows.length, data: rows };
}

function addInventory(p) {
  if (!p['Item Name']) return { ok: false, error: 'Item Name wajib diisi.' };
  p['Item ID'] = p['Item ID'] || generateId('P');
  appendToSheet(SH_INVENTORY, COLS_INVENTORY, p);
  return { ok: true, id: p['Item ID'] };
}

function updateInventory(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  updateSheetRow(SH_INVENTORY, COLS_INVENTORY, p._row, p);
  return { ok: true };
}

function deleteInventory(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  // Proteksi: jangan hapus jika Remaining QTY > 0
  const rows = rangeToObjects(RNG_INVENTORY);
  const item = rows.find(r => r._row === Number(p._row));
  if (item && parseNum(item['Remaining QTY']) > 0) {
    return { ok: false, error: 'Tidak bisa hapus — masih ada stok (Remaining QTY > 0).' };
  }
  deleteSheetRow(SH_INVENTORY, Number(p._row));
  return { ok: true };
}

// ── DIMENSIONS ───────────────────────────────────────────────────────────────
function getDimensions(p) {
  const rows = rangeToObjects(RNG_DIMENSIONS);
  // Return sebagai object map: { 'Item Type': ['Electronics', ...], ... }
  const map = {};
  rows.forEach(r => {
    const col = (r['Column Name'] || '').toString().trim();
    const val = (r['Value'] || '').toString().trim();
    if (!col || !val) return;
    map[col] = map[col] || [];
    map[col].push(val);
  });
  return { ok: true, data: map };
}

function addDimension(p) {
  if (!p['Column Name'] || !p['Value']) return { ok: false, error: 'Column Name dan Value wajib diisi.' };
  // Cek duplikat
  const rows = rangeToObjects(RNG_DIMENSIONS);
  const dup = rows.find(r =>
    r['Column Name'] === p['Column Name'] && r['Value'] === p['Value']
  );
  if (dup) return { ok: false, error: 'Nilai sudah ada.' };
  appendToSheet(SH_DIMENSIONS, COLS_DIMENSIONS, p);
  return { ok: true };
}

function deleteDimension(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  deleteSheetRow(SH_DIMENSIONS, Number(p._row));
  return { ok: true };
}

// ── SUPPLIERS ────────────────────────────────────────────────────────────────
function getSuppliers(p) {
  let rows = rangeToObjects(RNG_SUPPLIERS);
  if (p.state) rows = rows.filter(r => r['State'] === p.state);
  const NUM = ['Total Purchases','Total Payments','Balance Payable'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addSupplier(p) {
  if (!p['Supplier Name']) return { ok: false, error: 'Supplier Name wajib diisi.' };
  p['Supplier ID'] = p['Supplier ID'] || generateId('S');
  p['Total Purchases'] = p['Total Purchases'] || 0;
  p['Total Payments']  = p['Total Payments']  || 0;
  p['Balance Payable'] = p['Balance Payable'] || 0;
  appendToSheet(SH_SUPPLIERS, COLS_SUPPLIERS, p);
  return { ok: true, id: p['Supplier ID'] };
}

function updateSupplier(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  updateSheetRow(SH_SUPPLIERS, COLS_SUPPLIERS, p._row, p);
  return { ok: true };
}

function deleteSupplier(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const rows = rangeToObjects(RNG_SUPPLIERS);
  const item = rows.find(r => r._row === Number(p._row));
  if (item && parseNum(item['Balance Payable']) > 0) {
    return { ok: false, error: 'Tidak bisa hapus — masih ada Balance Payable.' };
  }
  deleteSheetRow(SH_SUPPLIERS, Number(p._row));
  return { ok: true };
}

// ── CUSTOMERS ─────────────────────────────────────────────────────────────────
function getCustomers(p) {
  let rows = rangeToObjects(RNG_CUSTOMERS);
  if (p.state) rows = rows.filter(r => r['State'] === p.state);
  const NUM = ['Total Sales','Total Receipts','Balance Receivable'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addCustomer(p) {
  if (!p['Customer Name']) return { ok: false, error: 'Customer Name wajib diisi.' };
  p['Customer ID'] = p['Customer ID'] || generateId('C');
  p['Total Sales']   = p['Total Sales']   || 0;
  p['Total Receipts'] = p['Total Receipts'] || 0;
  p['Balance Receivable'] = p['Balance Receivable'] || 0;
  appendToSheet(SH_CUSTOMERS, COLS_CUSTOMERS, p);
  return { ok: true, id: p['Customer ID'] };
}

function updateCustomer(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  updateSheetRow(SH_CUSTOMERS, COLS_CUSTOMERS, p._row, p);
  return { ok: true };
}

function deleteCustomer(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const rows = rangeToObjects(RNG_CUSTOMERS);
  const item = rows.find(r => r._row === Number(p._row));
  if (item && parseNum(item['Balance Receivable']) > 0) {
    return { ok: false, error: 'Tidak bisa hapus — masih ada Balance Receivable.' };
  }
  deleteSheetRow(SH_CUSTOMERS, Number(p._row));
  return { ok: true };
}

// ── PURCHASE ORDERS ──────────────────────────────────────────────────────────
function getPurchaseOrders(p) {
  let rows = rangeToObjects(RNG_PO);
  if (p.supplierId)   rows = rows.filter(r => r['Supplier ID'] === p.supplierId);
  if (p.pmtStatus)    rows = rows.filter(r => r['PMT Status'] === p.pmtStatus);
  if (p.shippingStatus) rows = rows.filter(r => r['Shipping Status'] === p.shippingStatus);
  const NUM = ['Total Amount','Total Paid','PO Balance'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addPurchaseOrder(p) {
  if (!p['Supplier Name']) return { ok: false, error: 'Supplier Name wajib diisi.' };
  p['PO ID'] = p['PO ID'] || generateId('PO');
  const total = parseNum(p['Total Amount']);
  const paid  = parseNum(p['Total Paid']);
  p['PO Balance'] = total - paid;
  if (!p['PMT Status']) p['PMT Status'] = paid >= total ? 'Paid' : paid > 0 ? 'Partial' : 'Unpaid';
  appendToSheet(SH_PO, COLS_PO, p);
  return { ok: true, id: p['PO ID'] };
}

function updatePurchaseOrder(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const total = parseNum(p['Total Amount']);
  const paid  = parseNum(p['Total Paid']);
  p['PO Balance'] = total - paid;
  updateSheetRow(SH_PO, COLS_PO, p._row, p);
  return { ok: true };
}

function deletePurchaseOrder(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const poId = p['PO ID'] || p.poId;
  // Hapus PO
  deleteSheetRow(SH_PO, Number(p._row));
  // Hapus semua Purchase Details dengan PO ID yang sama
  if (poId) {
    let pdRows = rangeToObjects(RNG_PD);
    pdRows = pdRows.filter(r => r['PO ID'] === poId);
    // Hapus dari bawah ke atas agar _row tidak bergeser
    pdRows.sort((a, b) => b._row - a._row);
    pdRows.forEach(r => deleteSheetRow(SH_PD, r._row));
  }
  return { ok: true };
}

// ── PURCHASE DETAILS ─────────────────────────────────────────────────────────
function getPurchaseDetails(p) {
  let rows = rangeToObjects(RNG_PD);
  if (p.poId)       rows = rows.filter(r => r['PO ID'] === p.poId);
  if (p.supplierId) rows = rows.filter(r => r['Supplier ID'] === p.supplierId);
  if (p.itemId)     rows = rows.filter(r => r['Item ID'] === p.itemId);
  const NUM = ['QTY Purchased','Unit Price','Total Purchase Price'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addPurchaseDetail(p) {
  if (!p['PO ID'] || !p['Item ID']) return { ok: false, error: 'PO ID dan Item ID wajib diisi.' };
  const qty   = parseNum(p['QTY Purchased']);
  const price = parseNum(p['Unit Price']);
  p['Total Purchase Price'] = qty * price;
  appendToSheet(SH_PD, COLS_PD, p);
  return { ok: true };
}

function deletePurchaseDetail(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  deleteSheetRow(SH_PD, Number(p._row));
  return { ok: true };
}

// ── SALES ORDERS ─────────────────────────────────────────────────────────────
function getSalesOrders(p) {
  let rows = rangeToObjects(RNG_SO);
  if (p.customerId)    rows = rows.filter(r => r['Customer ID'] === p.customerId);
  if (p.receiptStatus) rows = rows.filter(r => r['Receipt Status'] === p.receiptStatus);
  const NUM = ['Total SO Amount','Total Received','SO Balance'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addSalesOrder(p) {
  if (!p['Customer Name']) return { ok: false, error: 'Customer Name wajib diisi.' };
  p['SO ID'] = p['SO ID'] || generateId('SO');
  const total    = parseNum(p['Total SO Amount']);
  const received = parseNum(p['Total Received'] || 0);
  p['SO Balance'] = total - received;
  if (!p['Receipt Status']) p['Receipt Status'] = received >= total ? 'Received' : received > 0 ? 'Partial Receipt' : 'Pending';
  appendToSheet(SH_SO, COLS_SO, p);
  return { ok: true, id: p['SO ID'] };
}

function updateSalesOrder(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const total    = parseNum(p['Total SO Amount']);
  const received = parseNum(p['Total Received'] || 0);
  p['SO Balance'] = total - received;
  updateSheetRow(SH_SO, COLS_SO, p._row, p);
  return { ok: true };
}

function deleteSalesOrder(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const soId = p['SO ID'] || p.soId;
  deleteSheetRow(SH_SO, Number(p._row));
  // Hapus semua Sales Details dengan SO ID yang sama
  if (soId) {
    let sdRows = rangeToObjects(RNG_SD);
    sdRows = sdRows.filter(r => r['SO ID'] === soId);
    sdRows.sort((a, b) => b._row - a._row);
    sdRows.forEach(r => deleteSheetRow(SH_SD, r._row));
  }
  return { ok: true };
}

// ── SALES DETAILS ────────────────────────────────────────────────────────────
function getSalesDetails(p) {
  let rows = rangeToObjects(RNG_SD);
  if (p.soId)       rows = rows.filter(r => r['SO ID'] === p.soId);
  if (p.customerId) rows = rows.filter(r => r['Customer ID'] === p.customerId);
  if (p.itemId)     rows = rows.filter(r => r['Item ID'] === p.itemId);
  const NUM = ['QTY Sold','Unit Price','Price Incl Tax','Shipping Fees','Total Sales Price'];
  rows = rows.map(r => { const o = Object.assign({}, r); NUM.forEach(k => { o[k] = parseNum(o[k]); }); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addSalesDetail(p) {
  if (!p['SO ID'] || !p['Item ID']) return { ok: false, error: 'SO ID dan Item ID wajib diisi.' };
  p['Detail ID'] = p['Detail ID'] || generateId('SD');
  const qty    = parseNum(p['QTY Sold']);
  const price  = parseNum(p['Unit Price']);
  const inclTax= parseNum(p['Price Incl Tax'] || price * 1.1);
  const ship   = parseNum(p['Shipping Fees'] || 0);
  p['Price Incl Tax']    = inclTax;
  p['Total Sales Price'] = qty * inclTax + ship;
  appendToSheet(SH_SD, COLS_SD, p);
  return { ok: true, id: p['Detail ID'] };
}

function deleteSalesDetail(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  deleteSheetRow(SH_SD, Number(p._row));
  return { ok: true };
}

// ── RECEIPTS ─────────────────────────────────────────────────────────────────
function getReceipts(p) {
  let rows = rangeToObjects(RNG_RECEIPTS);
  if (p.customerId) rows = rows.filter(r => r['Customer ID'] === p.customerId);
  if (p.soId)       rows = rows.filter(r => r['SO ID'] === p.soId);
  if (p.pmtMode)    rows = rows.filter(r => r['PMT Mode'] === p.pmtMode);
  rows = rows.map(r => { const o = Object.assign({}, r); o['Amount Received'] = parseNum(o['Amount Received']); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addReceipt(p) {
  if (!p['Customer ID']) return { ok: false, error: 'Customer ID wajib diisi.' };
  p['Trx ID'] = p['Trx ID'] || generateId('RT');
  p['Amount Received'] = parseNum(p['Amount Received'] || 0);
  appendToSheet(SH_RECEIPTS, COLS_RECEIPTS, p);
  return { ok: true, id: p['Trx ID'] };
}

function updateReceipt(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  updateSheetRow(SH_RECEIPTS, COLS_RECEIPTS, p._row, p);
  return { ok: true };
}

function deleteReceipt(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  deleteSheetRow(SH_RECEIPTS, Number(p._row));
  return { ok: true };
}

// ── PAYMENTS ─────────────────────────────────────────────────────────────────
function getPayments(p) {
  let rows = rangeToObjects(RNG_PAYMENTS);
  if (p.supplierId) rows = rows.filter(r => r['Supplier ID'] === p.supplierId);
  if (p.poId)       rows = rows.filter(r => r['PO ID'] === p.poId);
  if (p.pmtMode)    rows = rows.filter(r => r['PMT Mode'] === p.pmtMode);
  rows = rows.map(r => { const o = Object.assign({}, r); o['Amount Paid'] = parseNum(o['Amount Paid']); return o; });
  return { ok: true, count: rows.length, data: rows };
}

function addPayment(p) {
  if (!p['Supplier ID']) return { ok: false, error: 'Supplier ID wajib diisi.' };
  p['Trx ID'] = p['Trx ID'] || generateId('PT');
  p['Amount Paid'] = parseNum(p['Amount Paid'] || 0);
  appendToSheet(SH_PAYMENTS, COLS_PAYMENTS, p);
  return { ok: true, id: p['Trx ID'] };
}

function updatePayment(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  updateSheetRow(SH_PAYMENTS, COLS_PAYMENTS, p._row, p);
  return { ok: true };
}

function deletePayment(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  deleteSheetRow(SH_PAYMENTS, Number(p._row));
  return { ok: true };
}

// ── USERS (Settings) ─────────────────────────────────────────────────────────
function getUsers(p) {
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, count: 0, data: [] };

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();

  const iHash = headers.indexOf('Password Hash');

  const rows = data
    .map((row, i) => {
      const obj = { _row: i + 2 };
      headers.forEach((h, j) => {
        if (h !== 'Password Hash') obj[h] = row[j] !== undefined ? row[j] : '';
      });
      return obj;
    })
    // Filter baris kosong SETELAH _row dihitung, supaya nomor baris tetap akurat
    // walau ada baris kosong di tengah (misalnya bekas dihapus manual, bukan lewat "Delete row").
    .filter(obj => Object.keys(obj).some(k => k !== '_row' && obj[k] !== ''));

  return { ok: true, count: rows.length, data: rows };
}

function addUser(p) {
  if (!p.Username || !p.Name || !p.Password) {
    return { ok: false, error: 'Username, Name, dan Password wajib diisi.' };
  }
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) return { ok: false, error: 'Sheet Users tidak ditemukan.' };

  // Cek duplikat username
  if (sh.getLastRow() >= 2) {
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const iUser   = headers.indexOf('Username');
    const dup     = data.find(row => (row[iUser] || '').toString().trim().toLowerCase() === p.Username.toLowerCase());
    if (dup) return { ok: false, error: 'Username "' + p.Username + '" sudah ada.' };
  }

  const userId = generateId('USR-');
  const hash   = hashPassword(p.Password);
  const row = [
    userId, p.Name, p.Email || '',
    p.Role || 'Staff', p.Status || 'Active',
    new Date().toISOString().substring(0, 10),
    p.Username, hash, p['Linked Customer ID'] || ''
  ];
  sh.appendRow(row);
  return { ok: true, id: userId };
}

function updateUser(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) return { ok: false, error: 'Sheet Users tidak ditemukan.' };

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const rowData = sh.getRange(p._row, 1, 1, lastCol).getValues()[0];

  // Update field yang dikirim (kecuali Password Hash langsung)
  const updateable = ['Name','Email','Role','Status','Linked Customer ID'];
  const iUser = headers.indexOf('Username');
  updateable.forEach(field => {
    const idx = headers.indexOf(field);
    if (idx > -1 && p[field] !== undefined) rowData[idx] = p[field];
  });
  // Username boleh diupdate jika tidak duplikat
  if (p.Username) {
    const data = sh.getLastRow() > 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues() : [];
    const dup  = data.find((row, i) => (i + 2) !== p._row &&
      (row[iUser] || '').toString().trim().toLowerCase() === p.Username.toLowerCase());
    if (dup) return { ok: false, error: 'Username "' + p.Username + '" sudah dipakai.' };
    rowData[iUser] = p.Username;
  }

  sh.getRange(p._row, 1, 1, lastCol).setValues([rowData]);
  return { ok: true };
}

function setUserPassword(p) {
  if (!p._row || !p.Password) return { ok: false, error: '_row dan Password wajib diisi.' };
  if (p.Password.length < 4) return { ok: false, error: 'Password minimal 4 karakter.' };
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) return { ok: false, error: 'Sheet Users tidak ditemukan.' };

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const iHash   = headers.indexOf('Password Hash');
  if (iHash < 0) return { ok: false, error: 'Kolom Password Hash tidak ditemukan.' };

  sh.getRange(p._row, iHash + 1).setValue(hashPassword(p.Password));
  return { ok: true };
}

function deleteUser(p) {
  if (!p._row) return { ok: false, error: '_row wajib diisi.' };
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) return { ok: false, error: 'Sheet Users tidak ditemukan.' };
  // Cegah hapus diri sendiri (token user aktif)
  if (p.token) {
    const session = CacheService.getScriptCache().get('ks_session_' + p.token);
    if (session) {
      const u = JSON.parse(session);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const row     = sh.getRange(p._row, 1, 1, sh.getLastColumn()).getValues()[0];
      const iUser   = headers.indexOf('Username');
      if (row[iUser] === u.username) return { ok: false, error: 'Tidak bisa menghapus akun yang sedang login.' };
    }
  }
  sh.deleteRow(Number(p._row));
  return { ok: true };
}

// ── SEED DEFAULT ADMIN ───────────────────────────────────────────────────────
/**
 * Jalankan dari Editor jika perlu reset/buat ulang user admin.
 */
function seedDefaultAdmin() {
  const ss = getSS();
  ensureSheet(ss, SH_USERS, COLS_USERS);
  const sh = ss.getSheetByName(SH_USERS);

  if (sh.getLastRow() >= 2) {
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
    const iUser   = headers.indexOf('Username');
    const iHash   = headers.indexOf('Password Hash');
    const iStatus = headers.indexOf('Status');
    for (let i = 0; i < data.length; i++) {
      if ((data[i][iUser] || '').toString().trim().toLowerCase() === 'admin') {
        sh.getRange(i + 2, iHash + 1).setValue(hashPassword('admin123'));
        sh.getRange(i + 2, iStatus + 1).setValue('Active');
        Logger.log('✅ Password admin di-reset ke "admin123".');
        SpreadsheetApp.getUi().alert('✅ Password admin di-reset ke "admin123".\nSilakan login dan ganti password.');
        return;
      }
    }
  }

  // Buat baru
  const userId = generateId('USR-');
  sh.appendRow([userId, 'Administrator', 'admin@kopsus.id', 'Admin', 'Active',
    new Date().toISOString().substring(0, 10), 'admin', hashPassword('admin123'), '']);
  Logger.log('✅ User admin dibuat. Username: admin / Password: admin123');
  SpreadsheetApp.getUi().alert('✅ User admin berhasil dibuat!\n\nUsername: admin\nPassword: admin123\n\nSegera ganti password setelah login.');
}

/**
 * Debug: tampilkan semua user di Logger (tanpa password hash).
 */
function debugListUsers() {
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) { Logger.log('Sheet Users kosong.'); return; }
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  const iName = headers.indexOf('Name'), iRole = headers.indexOf('Role');
  const iUser = headers.indexOf('Username'), iHash = headers.indexOf('Password Hash');
  Logger.log('=== Users ===');
  data.forEach((row, i) => {
    Logger.log('Row ' + (i+2) + ': ' + row[iName] + ' | ' + row[iRole] + ' | @' + row[iUser] + ' | hash: ' + ((row[iHash]||'').length > 0 ? '✅' : '❌'));
  });
}

/**
 * Migrasi: hash ulang password plain text yang tersimpan di sheet.
 * Jalankan sekali jika ada user yang passwordnya belum di-hash.
 */
function migratePasswords() {
  const ss = getSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh || sh.getLastRow() < 2) { Logger.log('Sheet Users kosong.'); return; }
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const data    = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  const iHash   = headers.indexOf('Password Hash');
  const iUser   = headers.indexOf('Username');
  if (iHash < 0) { Logger.log('Kolom Password Hash tidak ditemukan.'); return; }
  const hexRe   = /^[0-9a-f]{64}$/i;
  let fixed = 0;
  data.forEach((row, i) => {
    const stored = (row[iHash] || '').toString().trim();
    if (stored && !hexRe.test(stored)) {
      sh.getRange(i + 2, iHash + 1).setValue(hashPassword(stored));
      Logger.log('✅ Row ' + (i+2) + ' (@' + row[iUser] + '): password di-hash ulang.');
      fixed++;
    }
  });
  const msg = fixed + ' password berhasil di-hash ulang.';
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}
