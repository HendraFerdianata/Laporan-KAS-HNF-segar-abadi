/**
 * BUKU KAS - Backend API (Google Apps Script)
 * ---------------------------------------------
 * Script ini menjadikan Google Sheets sebagai "database" untuk dashboard
 * keuangan operasional. Menyediakan endpoint untuk: list, create, update,
 * delete transaksi.
 *
 * CARA PASANG:
 * 1. Buka Google Sheets baru, beri nama sheet pertama "Transaksi".
 * 2. Di baris 1, isi header persis seperti ini (kolom A-G):
 *    ID | Tanggal | Nama | Kategori | Masuk | Keluar | Catatan
 * 3. Buka menu Extensions > Apps Script, hapus isi default, lalu tempel
 *    seluruh isi file ini.
 * 4. Klik Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Salin URL Web App yang diberikan, itu adalah API_URL untuk dashboard.
 */

const SHEET_NAME = 'Transaksi';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ID', 'Tanggal', 'Nama', 'Kategori', 'Masuk', 'Keluar', 'Catatan']);
  }
  return sheet;
}

function jsonOut_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Mengambil seluruh transaksi sebagai array of object */
function getAllTransactions_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // buang header
  return rows
    .filter(r => r[0] !== '' && r[0] !== null) // buang baris kosong
    .map(r => ({
      id: r[0],
      tanggal: formatDate_(r[1]),
      nama: r[2],
      kategori: r[3],
      masuk: Number(r[4]) || 0,
      keluar: Number(r[5]) || 0,
      catatan: r[6] || ''
    }));
}

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

function findRowById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return i + 1; // nomor baris di sheet (1-indexed)
    }
  }
  return -1;
}

/** Tambah transaksi baru */
function createTransaction_(data) {
  const sheet = getSheet_();
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    data.tanggal,
    data.nama,
    data.kategori || '',
    Number(data.masuk) || 0,
    Number(data.keluar) || 0,
    data.catatan || ''
  ]);
  return { success: true, id: id };
}

/** Ubah transaksi berdasarkan ID */
function updateTransaction_(data) {
  const sheet = getSheet_();
  const row = findRowById_(sheet, data.id);
  if (row === -1) return { success: false, error: 'Transaksi tidak ditemukan' };

  sheet.getRange(row, 2, 1, 6).setValues([[
    data.tanggal,
    data.nama,
    data.kategori || '',
    Number(data.masuk) || 0,
    Number(data.keluar) || 0,
    data.catatan || ''
  ]]);
  return { success: true };
}

/** Hapus transaksi berdasarkan ID */
function deleteTransaction_(id) {
  const sheet = getSheet_();
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, error: 'Transaksi tidak ditemukan' };

  sheet.deleteRow(row);
  return { success: true };
}

/** Endpoint GET - dipakai untuk mengambil semua data (real-time saat dashboard dibuka) */
function doGet(e) {
  try {
    const action = (e.parameter.action || 'list');
    if (action === 'list') {
      return jsonOut_({ success: true, data: getAllTransactions_() });
    }
    return jsonOut_({ success: false, error: 'Action tidak dikenali' });
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

/** Endpoint POST - dipakai untuk create / update / delete */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if (action === 'create') {
      result = createTransaction_(body.data);
    } else if (action === 'update') {
      result = updateTransaction_(body.data);
    } else if (action === 'delete') {
      result = deleteTransaction_(body.id);
    } else {
      result = { success: false, error: 'Action tidak dikenali' };
    }

    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}
