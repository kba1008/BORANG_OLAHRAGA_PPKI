const SPREADSHEET_ID = "1rcfHlyjMoEVJw1cv6RZvo8vXT2vuURTHAVqxT6McjQw";
const FOLDER_ID = "1Nz0S__dRbA4vP4Ca0xBRhpdPNUj4KVOf";
const SHEET_PENGGUNA = "Pengguna";
const SHEET_SIJIL = "SijilPelajar";
const SHEET_PELAJAR = "SenaraiPelajar";

function doPost(e) {
  var output = ContentService.createTextOutput(); output.setMimeType(ContentService.MimeType.JSON);
  try {
    var data = JSON.parse(e.postData.contents); var action = data.action;

    if (action === "register") return output.setContent(JSON.stringify(daftarPengguna(data.email, data.password))); 
    else if (action === "login") return output.setContent(JSON.stringify(semakLogin(data.email, data.password))); 
    else if (action === "upload") return output.setContent(JSON.stringify(muatNaikSijil(data))); 
    else if (action === "updateFile") return output.setContent(JSON.stringify(kemaskiniSijil(data))); 
    else if (action === "getStudents") return output.setContent(JSON.stringify(dapatkanPelajar())); 
    else if (action === "addStudent") return output.setContent(JSON.stringify(tambahPelajar(data))); 
    else if (action === "getCerts") return output.setContent(JSON.stringify(dapatkanSijil(data.studentName))); 
    else if (action === "getRawPdf") return output.setContent(JSON.stringify(tarikPdfUntukDiedit(data.fileUrl)));
    
    // FUNGSI BARU: KEMASKINI MAKLUMAT PELAJAR
    else if (action === "updateStudent") return output.setContent(JSON.stringify(kemaskiniMaklumatPelajar(data)));

  } catch (err) { return output.setContent(JSON.stringify({ status: "error", message: err.toString() })); }
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName); var headers = [];
    if (sheetName === SHEET_PENGGUNA) headers = ["Email", "Kata Laluan", "Peranan", "Status"];
    else if (sheetName === SHEET_SIJIL) headers = ["Tarikh & Masa", "Nama Pelajar", "Nama Sijil", "Email Guru", "Pautan Fail Drive (PDF)"];
    else if (sheetName === SHEET_PELAJAR) headers = ["Nama Pelajar", "Kelas", "No IC", "Jantina", "Kategori", "Gambar Profil (Base64)", "Didaftar Oleh", "Tarikh Daftar"];
    if (headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
      sheet.setFrozenRows(1); for (var col = 1; col <= headers.length; col++) sheet.autoResizeColumn(col);
    }
  } return sheet;
}

function daftarPengguna(email, password) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_PENGGUNA);
  sheet.appendRow([email, password, (email === "admin") ? "Admin" : "User", (email === "admin") ? "Approved" : "Pending"]);
  return { status: "success", message: "Pendaftaran berjaya. Sila tunggu kelulusan Admin." };
}

function semakLogin(email, password) {
  if (email === "admin" && password === "101010") return { status: "success", role: "Admin" };
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_PENGGUNA), records = sheet.getDataRange().getValues();
  for (var i = 1; i < records.length; i++) {
    if (records[i][0] == email && records[i][1] == password) {
      if (records[i][3] == "Approved") return { status: "success", role: "User" };
      return { status: "error", message: "Akaun masih Pending." };
    }
  } return { status: "error", message: "Email atau Kata laluan salah." };
}

function muatNaikSijil(data) {
  var pdfBase64 = data.pdfData.split(',')[1], blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', data.fileName + ".pdf");
  var fileUrl = DriveApp.getFolderById(FOLDER_ID).createFile(blob).getUrl();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_SIJIL);
  sheet.appendRow([Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm"), data.studentName, data.certName, data.guruEmail, fileUrl]);
  return { status: "success", message: "Berjaya!", url: fileUrl };
}

function kemaskiniSijil(data) {
  var pdfBase64 = data.pdfData.split(',')[1], blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', data.fileName + "_Kemaskini.pdf");
  var fileUrl = DriveApp.getFolderById(FOLDER_ID).createFile(blob).getUrl();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_SIJIL), records = sheet.getDataRange().getValues();
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  for (var i = 1; i < records.length; i++) {
    if (records[i][1] == data.studentName && records[i][2] == data.certName) {
      sheet.getRange(i + 1, 1).setValue(tarikh); sheet.getRange(i + 1, 5).setValue(fileUrl);
      return { status: "success", message: "Fail dikemas kini!", url: fileUrl };
    }
  }
  sheet.appendRow([tarikh, data.studentName, data.certName, data.guruEmail, fileUrl]);
  return { status: "success", message: "Rekod baru disimpan.", url: fileUrl };
}

function dapatkanPelajar() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var data = sheet.getDataRange().getValues(), students = [];
  for (var i = 1; i < data.length; i++) { 
    if (data[i][0]) {
      students.push({
        name: data[i][0].toString().trim(), kelas: data[i][1] || "", ic: data[i][2] || "",
        jantina: data[i][3] || "", kategori: data[i][4] || "", gambar: data[i][5] || ""
      });
    }
  }
  students.sort((a, b) => a.name.localeCompare(b.name)); 
  return { status: "success", data: students };
}

function tambahPelajar(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var records = sheet.getDataRange().getValues();
  var nameTrimmed = data.studentName.trim();
  for (var i = 1; i < records.length; i++) {
    if (records[i][0].toString().trim().toLowerCase() === nameTrimmed.toLowerCase()) {
      return { status: "error", message: "Nama pelajar ini sudah direkodkan di dalam sistem!" };
    }
  }
  var picData = data.gambarBase64 ? data.gambarBase64 : "";
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  sheet.appendRow([nameTrimmed, data.kelas || "-", data.ic || "-", data.jantina || "-", data.kategori || "-", picData, data.guruEmail, tarikh]);
  return { status: "success", message: "Pelajar berjaya didaftarkan." };
}

// FUNGSI BARU: UPDATE PROFIL PELAJAR
function kemaskiniMaklumatPelajar(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var records = sheet.getDataRange().getValues();
  var oldName = data.oldName.trim();
  var newName = data.studentName.trim();

  // Jika nama ditukar, pastikan nama baru tak bertindih dengan orang lain
  if(oldName.toLowerCase() !== newName.toLowerCase()) {
     for (var j = 1; j < records.length; j++) {
        if (records[j][0].toString().trim().toLowerCase() === newName.toLowerCase()) {
          return { status: "error", message: "Gagal: Nama baru ini sudah wujud dalam sistem." };
        }
     }
  }

  // Cari dan Update
  for (var i = 1; i < records.length; i++) {
    if (records[i][0].toString().trim() === oldName) {
      sheet.getRange(i + 1, 1).setValue(newName);
      sheet.getRange(i + 1, 2).setValue(data.kelas || "-");
      sheet.getRange(i + 1, 3).setValue(data.ic || "-");
      sheet.getRange(i + 1, 4).setValue(data.jantina || "-");
      sheet.getRange(i + 1, 5).setValue(data.kategori || "-");
      
      // Jika ada gambar baru diupload, update. Kalau tak, kekalkan yang lama.
      if (data.gambarBase64) {
         sheet.getRange(i + 1, 6).setValue(data.gambarBase64);
      }
      
      // Jika nama bertukar, update rekod SijilPelajar juga supaya sijil lama tak hilang
      if(oldName !== newName) {
         var sheetSijil = getOrCreateSheet(ss, SHEET_SIJIL);
         var certRecords = sheetSijil.getDataRange().getValues();
         for(var k=1; k<certRecords.length; k++) {
            if(certRecords[k][1].toString().trim() === oldName) {
                sheetSijil.getRange(k+1, 2).setValue(newName);
            }
         }
      }
      return { status: "success", message: "Profil pelajar berjaya dikemas kini!" };
    }
  }
  return { status: "error", message: "Pelajar tidak dijumpai dalam pangkalan data." };
}

function dapatkanSijil(studentName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_SIJIL);
  var records = sheet.getDataRange().getValues();
  var certs = [];
  for (var i = 1; i < records.length; i++) {
     if (records[i][1] === studentName) {
        certs.push({ date: records[i][0], certName: records[i][2], url: records[i][4] });
     }
  }
  certs.reverse();
  return { status: "success", data: certs };
}

function tarikPdfUntukDiedit(url) {
  try {
    var idMatch = url.match(/[-\w]{25,}/);
    if (!idMatch) return { status: "error", message: "ID fail tidak sah." };
    var file = DriveApp.getFileById(idMatch[0]);
    var base64 = Utilities.base64Encode(file.getBlob().getBytes());
    return { status: "success", base64Data: base64 };
  } catch (err) {
    return { status: "error", message: "Gagal akses: " + err.toString() };
  }
}
