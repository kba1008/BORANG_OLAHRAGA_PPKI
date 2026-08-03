// ==============================================================
// TETAPAN GOOGLE API (SILA MASUKKAN ID ANDA DI SINI)
// ==============================================================
const SPREADSHEET_ID = "1rcfHlyjMoEVJw1cv6RZvo8vXT2vuURTHAVqxT6McjQw";
const FOLDER_ID = "1Nz0S__dRbA4vP4Ca0xBRhpdPNUj4KVOf";

const SHEET_PENGGUNA = "Pengguna";
const SHEET_SIJIL = "SijilPelajar";

// Fungsi utama yang menerima arahan dari Web/PWA
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    if (action === "register") {
      return output.setContent(JSON.stringify(daftarPengguna(data.email, data.password)));
    } 
    else if (action === "login") {
      return output.setContent(JSON.stringify(semakLogin(data.email, data.password)));
    } 
    else if (action === "upload") {
      return output.setContent(JSON.stringify(muatNaikSijil(data)));
    }
    
  } catch (err) {
    return output.setContent(JSON.stringify({ status: "error", message: err.toString() }));
  }
}

// FUNGSI 1: DAFTAR PENGGUNA
function daftarPengguna(email, password) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PENGGUNA);
  var role = (email === "admin") ? "Admin" : "User";
  var status = (email === "admin") ? "Approved" : "Pending";
  
  sheet.appendRow([email, password, role, status]);
  return { status: "success", message: "Pendaftaran berjaya. Sila tunggu kelulusan Admin." };
}

// FUNGSI 2: LOG MASUK
function semakLogin(email, password) {
  if (email === "admin" && password === "101010") return { status: "success", role: "Admin" };
  
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PENGGUNA);
  var records = sheet.getDataRange().getValues();
  
  for (var i = 1; i < records.length; i++) {
    if (records[i][0] == email && records[i][1] == password) {
      if (records[i][3] == "Approved") return { status: "success", role: "User" };
      return { status: "error", message: "Akaun anda masih Pending kelulusan." };
    }
  }
  return { status: "error", message: "Email atau Kata laluan salah." };
}

// FUNGSI 3: MUAT NAIK PDF KE DRIVE & REKOD KE SHEETS
function muatNaikSijil(data) {
  // Tukar Base64 kembali menjadi fail PDF
  var pdfBase64 = data.pdfData.split(',')[1]; 
  var blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', data.fileName + ".pdf");
  
  // Simpan ke Google Drive
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var file = folder.createFile(blob);
  var fileUrl = file.getUrl();
  
  // Rekod ke Google Sheet
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SIJIL);
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  
  sheet.appendRow([tarikh, data.studentName, data.certName, data.guruEmail, fileUrl]);
  
  return { status: "success", message: "Sijil berjaya dimuat naik ke Drive!", url: fileUrl };
}
