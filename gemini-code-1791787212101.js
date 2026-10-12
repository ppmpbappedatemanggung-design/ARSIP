/**
 * DOKUMENTASI SUBID PEMERINTAHAN PPMP
 * Aplikasi Penyimpanan Digital Terintegrasi Google Drive
 */

// Force OAuth Scopes agar Apps Script tidak meminta otorisasi berulang atau error DriveApp
// @OnlyCurrentDoc
// https://www.googleapis.com/auth/drive
// https://www.googleapis.com/auth/spreadsheets

const FOLDER_ID = '1GA0clpTf6E0Lc1g2ivGQzdrgrrQBp9VS';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DOKUMENTASI SUBID PEMERINTAHAN PPMP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fungsi Inisialisasi Database Spreadsheet
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const categories = [
    { name: 'TMMD', headers: ['ID', 'Timestamp', 'Judul', 'Tahun', 'Nomor SK', 'Nama File', 'File ID', 'URL View', 'URL Download'] },
    { name: 'RAT NANGKIS', headers: ['ID', 'Timestamp', 'Judul', 'Nama File', 'File ID', 'URL View', 'URL Download'] },
    { name: 'LP2KD', headers: ['ID', 'Timestamp', 'Judul', 'Nama File', 'File ID', 'URL View', 'URL Download'] },
    { name: 'LAPORAN OPPKPKE', headers: ['ID', 'Timestamp', 'Judul', 'Tahun', 'Semester', 'Nama File', 'File ID', 'URL View', 'URL Download'] },
    { name: 'SK PENGELOLA DTSEN', headers: ['ID', 'Timestamp', 'Judul', 'Nama File', 'File ID', 'URL View', 'URL Download'] },
    { name: 'KELENGKAPAN DTSEN', headers: ['ID', 'Timestamp', 'Judul', 'Tahun', 'Periode', 'Nama File', 'File ID', 'URL View', 'URL Download'] }
  ];

  categories.forEach(cat => {
    let sheet = ss.getSheetByName(cat.name);
    if (!sheet) {
      sheet = ss.insertSheet(cat.name);
      sheet.appendRow(cat.headers);
      sheet.getRange(1, 1, 1, cat.headers.length).setBackground('#1b365d').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  return 'Setup Database Berhasil!';
}

// Fungsi Mengunggah File ke Google Drive
function uploadFileToDrive(data) {
  try {
    let parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(FOLDER_ID);
    } catch (e) {
      return { 
        success: false, 
        message: 'Akses Drive ditolak atau ID Folder tidak ditemukan. Pastikan Anda telah mengizinkan akses Drive.' 
      };
    }

    // Cari/Buat Subfolder Kategori
    let subFolders = parentFolder.getFoldersByName(data.category);
    let targetFolder = subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(data.category);

    // Proses File Base64
    const base64Data = data.fileContent.substr(data.fileContent.indexOf(',') + 1);
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, data.mimeType, data.fileName);
    const file = targetFolder.createFile(blob);

    // Buka Akses Lihat File
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const fileUrl = file.getUrl();
    const downloadUrl = file.getDownloadUrl();
    const timestamp = new Date();
    const uniqueId = 'DOC-' + Math.floor(100000 + Math.random() * 900000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(data.category);

    if (!sheet) {
      throw new Error('Sheet ' + data.category + ' tidak ditemukan. Silakan jalankan setupDatabase terlebih dahulu.');
    }

    let rowData = [uniqueId, timestamp, data.judul];
    if (data.category === 'TMMD') {
      rowData.push(data.tahun || '', data.nomorSk || '');
    } else if (data.category === 'LAPORAN OPPKPKE') {
      rowData.push(data.tahun || '', data.semester || '');
    } else if (data.category === 'KELENGKAPAN DTSEN') {
      rowData.push(data.tahun || '', data.periode || '');
    }

    rowData.push(data.fileName, fileId, fileUrl, downloadUrl);
    sheet.appendRow(rowData);

    return { success: true, message: 'File berhasil diunggah dan disimpan ke Drive!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Mengambil Data Arsip untuk Dasbor
function getAllFilesData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const categories = ['TMMD', 'RAT NANGKIS', 'LP2KD', 'LAPORAN OPPKPKE', 'SK PENGELOLA DTSEN', 'KELENGKAPAN DTSEN'];
    let allData = {};

    categories.forEach(catName => {
      let sheet = ss.getSheetByName(catName);
      if (sheet) {
        let rows = sheet.getDataRange().getValues();
        if (rows.length > 1) {
          let headers = rows.shift();
          allData[catName] = rows.map(row => {
            let obj = {};
            headers.forEach((h, index) => {
              obj[h] = row[index];
            });
            return obj;
          });
        } else {
          allData[catName] = [];
        }
      } else {
        allData[catName] = [];
      }
    });

    return { success: true, data: allData };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// Fungsi Hapus Arsip
function deleteFileRecord(category, fileId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(category);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const fileIdIndex = headers.indexOf('File ID');

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][fileIdIndex] === fileId) {
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
        } catch (e) {
          console.warn('File tidak ditemukan di Drive atau sudah dihapus:', e);
        }
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true, message: 'Arsip berhasil dihapus.' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}