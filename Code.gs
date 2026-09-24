/**
 * ============================================================
 *  BACKEND PENDAFTARAN HAMKA - Google Apps Script
 * ============================================================
 * Cara pakai:
 * 1. Buka Google Sheet baru (kosong).
 * 2. Menu Extensions > Apps Script.
 * 3. Hapus isi default, tempel (paste) SELURUH kode file ini.
 * 4. Ganti FOLDER_ID di bawah dengan ID folder Google Drive
 *    tempat menyimpan file sertifikat pendaftar.
 * 5. Klik Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Salin URL Web App yang muncul, tempel ke variabel
 *    SCRIPT_URL di file HTML form kamu.
 * ============================================================
 */

// GANTI dengan ID folder Google Drive kamu (lihat panduan cara ambil ID folder)
const FOLDER_ID = '1lOy4HjfjeMQzfY-aEuUQTBF62GA-qWhV-UchyC';

// Nama sheet tempat data disimpan (akan dibuat otomatis kalau belum ada)
const SHEET_NAME = 'Pendaftaran';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Buat header kalau sheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Nama',
        'Kelas',
        'Program Studi',
        'Jenis Kelamin',
        'Alasan Mengikuti HAMKA',
        'Minat Departemen',
        'Sertifikat MORIEST',
        'Sertifikat LKM',
        'Sertifikat AMT'
      ]);
    }

    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Simpan 1 file base64 ke Drive, kembalikan URL file-nya
    function saveFile(fileObj, labelPrefix) {
      if (!fileObj || !fileObj.data) return '-';
      const bytes = Utilities.base64Decode(fileObj.data);
      const safeName = labelPrefix + '_' + (fileObj.fileName || 'file');
      const blob = Utilities.newBlob(bytes, fileObj.mimeType, safeName);
      const file = folder.createFile(blob);
      // Supaya panitia bisa buka linknya langsung tanpa minta akses satu-satu
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
    }

    const namaPendaftar = data.nama || 'Tanpa_Nama';
    const linkMoriest = saveFile(data.files && data.files.sertif_moriest, namaPendaftar + '_MORIEST');
    const linkLkm     = saveFile(data.files && data.files.sertif_lkm, namaPendaftar + '_LKM');
    const linkAmt      = saveFile(data.files && data.files.sertif_amt, namaPendaftar + '_AMT');

    sheet.appendRow([
      new Date(),
      data.nama || '',
      data.kelas || '',
      data.prodi || '',
      data.jenis_kelamin || '',
      data.alasan || '',
      (data.departemen || []).join(', '),
      linkMoriest,
      linkLkm,
      linkAmt
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Opsional: fungsi ini hanya untuk mengecek Web App aktif kalau dibuka lewat browser (GET)
function doGet(e) {
  return ContentService
    .createTextOutput('Backend Pendaftaran HAMKA aktif. Gunakan method POST untuk mengirim data.')
    .setMimeType(ContentService.MimeType.TEXT);
}
