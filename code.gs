// ==============================================================
// TETAPAN GOOGLE API 
// ==============================================================
const SPREADSHEET_ID = "1rcfHlyjMoEVJw1cv6RZvo8vXT2vuURTHAVqxT6McjQw";
const FOLDER_ID = "1Nz0S__dRbA4vP4Ca0xBRhpdPNUj4KVOf";

const SHEET_PENGGUNA = "Pengguna";
const SHEET_SIJIL = "SijilPelajar";

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

// ==============================================================
// FUNGSI BARU: AUTO-CREATE & SUSUN HEADER GOOGLE SHEET
// ==============================================================
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  
  // Jika tab/sheet tidak wujud, ia akan buat baru
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [];
    
    // Tetapkan nama kolum (header) berdasarkan nama sheet
    if (sheetName === SHEET_PENGGUNA) {
      headers = ["Email", "Kata Laluan", "Peranan", "Status"];
    } else if (sheetName === SHEET_SIJIL) {
      headers = ["Tarikh & Masa", "Nama Pelajar", "Nama Sijil", "Email Guru", "Pautan Fail Drive (PDF)"];
    }
    
    if (headers.length > 0) {
      // Masukkan baris header
      sheet.appendRow(headers);
      
      // Cantikkan jadual (Tulisan Tebal, Warna Latar, Teks Putih)
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4f46e5"); // Warna biru indigo (Sama dengan UI App)
      headerRange.setFontColor("white");
      
      // Bekukan (freeze) baris atas supaya tak bergerak bila skrol
      sheet.setFrozenRows(1);
      
      // Auto-adjust kelebaran setiap kolum supaya kemas
      for (var col = 1; col <= headers.length; col++) {
        sheet.autoResizeColumn(col);
      }
    }
  }
  return sheet;
}

// ==============================================================
// FUNGSI PENGKANGKALAN DATA (DATABASE)
// ==============================================================
function daftarPengguna(email, password) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PENGGUNA);
  
  var role = (email === "admin") ? "Admin" : "User";
  var status = (email === "admin") ? "Approved" : "Pending";
  
  sheet.appendRow([email, password, role, status]);
  return { status: "success", message: "Pendaftaran berjaya. Sila tunggu kelulusan Admin." };
}

function semakLogin(email, password) {
  if (email === "admin" && password === "101010") return { status: "success", role: "Admin" };
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PENGGUNA);
  var records = sheet.getDataRange().getValues();
  
  for (var i = 1; i < records.length; i++) {
    if (records[i][0] == email && records[i][1] == password) {
      if (records[i][3] == "Approved") return { status: "success", role: "User" };
      return { status: "error", message: "Akaun anda masih Pending kelulusan." };
    }
  }
  return { status: "error", message: "Email atau Kata laluan salah." };
}

function muatNaikSijil(data) {
  // Terima dan convert PDF Base64 kepada Blob
  var pdfBase64 = data.pdfData.split(',')[1]; 
  var blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', data.fileName + ".pdf");
  
  // Hantar ke Google Drive
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var file = folder.createFile(blob);
  var fileUrl = file.getUrl();
  
  // Rekod ke Google Sheets
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_SIJIL);
  
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  
  sheet.appendRow([tarikh, data.studentName, data.certName, data.guruEmail, fileUrl]);
  
  return { status: "success", message: "Sijil berjaya dimuat naik ke Drive!", url: fileUrl };
}
