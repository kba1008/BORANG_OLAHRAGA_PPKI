/**
 * SISTEM SIJIL PRO - BACKEND (Google Apps Script)
 * Versi 2.0 - Dioptimumkan untuk RIBUAN rekod
 *
 * PERUBAHAN UTAMA:
 * 1. Gambar profil disimpan sebagai FAIL DI DRIVE (bukan Base64 dalam Sheet).
 *    -> Sheet jadi ringan, loading jauh lebih laju.
 * 2. getStudents menyokong paging + carian di server (limit/offset).
 * 3. Baca lajur tertentu sahaja (bukan getDataRange penuh) = laju.
 * 4. CacheService untuk indeks nama pelajar.
 */

const SPREADSHEET_ID = "1rcfHlyjMoEVJw1cv6RZvo8vXT2vuURTHAVqxT6McjQw";
const FOLDER_ID = "1Nz0S__dRbA4vP4Ca0xBRhpdPNUj4KVOf";
const SHEET_PENGGUNA = "Pengguna";
const SHEET_SIJIL = "SijilPelajar";
const SHEET_PELAJAR = "SenaraiPelajar";
const FOLDER_GAMBAR_NAME = "Gambar Profil Pelajar";
const CACHE_SEC = 120;

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;

    if (action === "register") result = daftarPengguna(data);
    else if (action === "login") result = semakLogin(data.email, data.password);
    else if (action === "upload") result = muatNaikSijil(data);
    else if (action === "updateFile") result = kemaskiniSijil(data);
    else if (action === "getStudents") result = dapatkanPelajar(data);
    else if (action === "addStudent") result = tambahPelajar(data);
    else if (action === "getCerts") result = dapatkanSijil(data.studentName);
    else if (action === "getRawPdf") result = tarikPdfUntukDiedit(data.fileUrl);
    else if (action === "updateStudent") result = kemaskiniMaklumatPelajar(data);
    else if (action === "deleteStudent") result = buangPelajar(data.studentName);
    else if (action === "setSemakan") result = tetapSemakan(data);
    else if (action === "getUsers") result = senaraiPengguna(data);
    else if (action === "setUserStatus") result = tetapStatusPengguna(data);
    else if (action === "deleteUser") result = buangPengguna(data);
    else result = { status: "error", message: "Action tidak dikenali: " + action };


    return output.setContent(JSON.stringify(result));
  } catch (err) {
    return output.setContent(JSON.stringify({ status: "error", message: err.toString() }));
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Sijil Pro API v2 aktif." }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 * UTILITI
 * ============================================================ */
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [];
    if (sheetName === SHEET_PENGGUNA) headers = ["Email", "Kata Laluan", "Peranan", "Status", "Nama Penuh", "Nama Sekolah", "No Telefon", "Jawatan", "Tarikh Daftar"];
    else if (sheetName === SHEET_SIJIL) headers = ["Tarikh & Masa", "Nama Pelajar", "Nama Sijil", "Email Guru", "Pautan Fail Drive (PDF)", "Dikemaskini Oleh"];
    else if (sheetName === SHEET_PELAJAR) headers = ["Nama Pelajar", "Kelas", "No IC", "Jantina", "Kategori", "ID Gambar Drive", "Didaftar Oleh", "Tarikh Daftar", "Nama Sekolah"];
    if (headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getFolderGambar() {
  var parent = DriveApp.getFolderById(FOLDER_ID);
  var it = parent.getFoldersByName(FOLDER_GAMBAR_NAME);
  return it.hasNext() ? it.next() : parent.createFolder(FOLDER_GAMBAR_NAME);
}

/**
 * Simpan gambar profil ke Drive, pulangkan File ID.
 * Menerima data URI base64 dari browser (sudah dimampatkan ke ~256px).
 */
function simpanGambarProfil(base64DataUri, namaPelajar) {
  if (!base64DataUri) return "";
  try {
    var parts = base64DataUri.split(",");
    var meta = parts[0] || "";
    var raw = parts.length > 1 ? parts[1] : parts[0];
    var mime = meta.indexOf("png") > -1 ? "image/png" : "image/jpeg";
    var ext = mime === "image/png" ? ".png" : ".jpg";
    var blob = Utilities.newBlob(Utilities.base64Decode(raw), mime, "profil_" + namaPelajar.replace(/[^\w]/g, "_") + "_" + Date.now() + ext);
    var file = getFolderGambar().createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return file.getId();
  } catch (err) {
    return "";
  }
}

/** Baca hanya lajur ringan (nama/kelas/ic/jantina/kategori/idGambar) - tanpa base64 berat */
/** Pastikan lajur 9 = Nama Sekolah wujud (rekod lama hanya 8 lajur) */
function pastikanLajurSekolah(sheet) {
  if (sheet.getMaxColumns() < 9) sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());
  var head = sheet.getRange(1, 9).getValue();
  if (!head) sheet.getRange(1, 9).setValue("Nama Sekolah");
}

/** Pastikan lajur 10-13 (Semakan Borang) wujud */
function pastikanLajurSemakan(sheet) {
  if (sheet.getMaxColumns() < 13) sheet.insertColumnsAfter(sheet.getMaxColumns(), 13 - sheet.getMaxColumns());
  var tajuk = ["Status Semakan", "Catatan Semakan", "Disemak Oleh", "Tarikh Semakan"];
  var head = sheet.getRange(1, 10, 1, 4).getValues()[0];
  for (var i = 0; i < 4; i++) {
    if (!head[i]) sheet.getRange(1, 10 + i).setValue(tajuk[i]);
  }
}

function bacaSemuaPelajarRingkas(sheet) {
  pastikanLajurSekolah(sheet);
  pastikanLajurSemakan(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    var gambar = r[5] ? r[5].toString() : "";
    // Rekod lama menyimpan base64 penuh -> jangan hantar ke klien (berat!)
    if (gambar.indexOf("data:") === 0) gambar = "";
    out.push({
      row: i + 2,
      name: r[0].toString().trim(),
      kelas: r[1] || "",
      ic: r[2] ? r[2].toString() : "",
      jantina: r[3] || "",
      kategori: r[4] || "",
      sekolah: r[8] || "",
      gambarId: gambar,
      semakan: (r[9] || "").toString().trim(),
      catatan: (r[10] || "").toString(),
      semakOleh: (r[11] || "").toString(),
      semakTarikh: (r[12] || "").toString()
    });
  }
  return out;
}

/* ============================================================
 * AUTH
 * ============================================================ */
/** Pastikan akaun MASTER ADMIN wujud DALAM GOOGLE SHEET sahaja (bukan hardcode). */
function pastikanAdmin(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (emails[i][0] && emails[i][0].toString().trim().toLowerCase() === "admin") {
        // Pastikan peranan & status betul
        sheet.getRange(i + 2, 3, 1, 2).setValues([["Admin", "Approved"]]);
        return;
      }
    }
  }
  sheet.appendRow(["admin", "101010", "Admin", "Approved"]);
}

/** Pastikan lajur 5-9 (Nama Penuh, Sekolah, Telefon, Jawatan, Tarikh) wujud pada sheet lama */
function pastikanLajurPengguna(sheet) {
  if (sheet.getMaxColumns() < 9) sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());
  var tajuk = ["Email", "Kata Laluan", "Peranan", "Status", "Nama Penuh", "Nama Sekolah", "No Telefon", "Jawatan", "Tarikh Daftar"];
  var semasa = sheet.getRange(1, 1, 1, 9).getValues()[0];
  for (var i = 0; i < 9; i++) {
    if (!semasa[i]) sheet.getRange(1, i + 1).setValue(tajuk[i]);
  }
}

function sheetPengguna() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PENGGUNA);
  pastikanLajurPengguna(sheet);
  pastikanAdmin(sheet);
  return sheet;
}

function daftarPengguna(data) {
  var email = (data.email || "").toString().trim();
  var password = (data.password || "").toString();
  var nama = (data.namaPenuh || "").toString().trim();
  var sekolah = (data.sekolah || "").toString().trim();
  var telefon = (data.telefon || "").toString().trim();
  var jawatan = (data.jawatan || "").toString().trim();

  if (!nama) return { status: "error", message: "Nama penuh diperlukan." };
  if (!sekolah) return { status: "error", message: "Nama sekolah diperlukan." };
  if (!telefon) return { status: "error", message: "No telefon diperlukan." };
  if (!/^[0-9+\-\s]{8,20}$/.test(telefon)) return { status: "error", message: "No telefon tidak sah." };
  if (!jawatan) return { status: "error", message: "Jawatan diperlukan." };
  if (!email || !password) return { status: "error", message: "Emel dan kata laluan diperlukan." };
  if (password.length < 6) return { status: "error", message: "Kata laluan mesti sekurang-kurangnya 6 aksara." };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { status: "error", message: "Sila masukkan alamat emel yang sah." };

  var sheet = sheetPengguna();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (emails[i][0] && emails[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        return { status: "error", message: "Emel ini sudah didaftarkan." };
      }
    }
  }
  sheet.appendRow([email, password, "User", "Pending", nama, sekolah, telefon, jawatan,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")]);
  return { status: "success", message: "Pendaftaran berjaya. Akaun anda perlu diluluskan oleh Master Admin sebelum boleh digunakan." };
}

function semakLogin(email, password) {
  email = (email || "").toString().trim();
  password = (password || "").toString();
  var sheet = sheetPengguna();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "error", message: "Emel atau kata laluan salah." };
  var records = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (var i = 0; i < records.length; i++) {
    var e = records[i][0] ? records[i][0].toString().trim() : "";
    if (e.toLowerCase() !== email.toLowerCase()) continue;
    if (records[i][1].toString() !== password) return { status: "error", message: "Emel atau kata laluan salah." };
    var status = (records[i][3] || "").toString().trim();
    if (status === "Approved") return { status: "success", role: records[i][2] || "User", email: e, namaPenuh: records[i][4] || "", sekolah: records[i][5] || "" };
    if (status === "Rejected") return { status: "error", message: "Akaun anda telah ditolak oleh Admin." };
    return { status: "error", message: "Akaun anda masih menunggu kelulusan Master Admin." };
  }
  return { status: "error", message: "Emel atau kata laluan salah." };
}

/** Sahkan pemanggil ialah Master Admin (berdasarkan rekod dalam Sheet). */
function sahAdmin(email, password) {
  var res = semakLogin(email, password);
  return res.status === "success" && res.role === "Admin";
}

function senaraiPengguna(data) {
  if (!sahAdmin(data.adminEmail, data.adminPassword)) return { status: "error", message: "Akses ditolak. Admin sahaja." };
  var sheet = sheetPengguna();
  var lastRow = sheet.getLastRow();
  var out = [];
  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      out.push({
        row: i + 2, email: rows[i][0].toString(), password: rows[i][1] ? rows[i][1].toString() : "", role: rows[i][2] || "User", status: rows[i][3] || "Pending",
        namaPenuh: rows[i][4] || "", sekolah: rows[i][5] || "", telefon: rows[i][6] ? rows[i][6].toString() : "",
        jawatan: rows[i][7] || "", tarikh: rows[i][8] ? rows[i][8].toString() : ""
      });
    }
  }
  return { status: "success", data: out };
}

function tetapStatusPengguna(data) {
  if (!sahAdmin(data.adminEmail, data.adminPassword)) return { status: "error", message: "Akses ditolak. Admin sahaja." };
  var target = (data.email || "").toString().trim();
  if (target.toLowerCase() === "admin") return { status: "error", message: "Akaun Master Admin tidak boleh diubah." };
  var status = data.newStatus === "Approved" ? "Approved" : "Rejected";
  var sheet = sheetPengguna();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (emails[i][0] && emails[i][0].toString().trim().toLowerCase() === target.toLowerCase()) {
        sheet.getRange(i + 2, 4).setValue(status);
        return { status: "success", message: "Status " + target + " ditukar kepada " + status + "." };
      }
    }
  }
  return { status: "error", message: "Pengguna tidak dijumpai." };
}

function buangPengguna(data) {
  if (!sahAdmin(data.adminEmail, data.adminPassword)) return { status: "error", message: "Akses ditolak. Admin sahaja." };
  var target = (data.email || "").toString().trim();
  if (target.toLowerCase() === "admin") return { status: "error", message: "Akaun Master Admin tidak boleh dipadam." };
  var sheet = sheetPengguna();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (emails[i][0] && emails[i][0].toString().trim().toLowerCase() === target.toLowerCase()) {
        sheet.deleteRow(i + 2);
        return { status: "success", message: "Pengguna " + target + " dipadam." };
      }
    }
  }
  return { status: "error", message: "Pengguna tidak dijumpai." };
}


/* ============================================================
 * SIJIL (PDF)
 * ============================================================ */
function muatNaikSijil(data) {
  // Satu pelajar = satu fail terkini sahaja. Simpan dulu, padam lama kemudian.
  return kemaskiniSijil(data);
}


/** Padam (trash) fail PDF lama di Drive berdasarkan pautan. Selamat: tidak throw. */
function padamFailDrive(url, kecualiId) {
  try {
    if (!url) return false;
    var m = String(url).match(/[-\w]{25,}/);
    if (!m) return false;
    var id = m[0];
    if (kecualiId && id === kecualiId) return false;
    DriveApp.getFileById(id).setTrashed(true);
    return true;
  } catch (e) {
    return false;
  }
}

/** Pastikan sheet Sijil ada lajur "Dikemaskini Oleh" (lajur ke-6). */
function pastikanLajurSijil(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 6) {
    sheet.getRange(1, 6).setValue("Dikemaskini Oleh").setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
  }
}

function kemaskiniSijil(data) {
  // 1) SIMPAN DULU fail baharu. Jika internet/simpan gagal, fail lama kekal utuh.
  var pdfBase64 = data.pdfData.split(',')[1];
  var bytes = Utilities.base64Decode(pdfBase64);
  if (!bytes || bytes.length === 0) return { status: "error", message: "Data PDF kosong. Fail lama tidak disentuh." };

  var blob = Utilities.newBlob(bytes, 'application/pdf', data.fileName + ".pdf");
  var file = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var newId = file.getId();
  var fileUrl = file.getUrl();

  // 2) Sahkan fail baharu benar-benar wujud & bersaiz sebelum padam apa-apa.
  var okBaharu = false;
  try { okBaharu = DriveApp.getFileById(newId).getSize() > 0; } catch (e) { okBaharu = false; }
  if (!okBaharu) return { status: "error", message: "Fail baharu gagal disimpan di Drive. Fail lama tidak dipadam." };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_SIJIL);
  pastikanLajurSijil(sheet);
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  var dikemaskiniOleh = (data.guruNama || data.guruEmail || "").toString().trim();
  var lastRow = sheet.getLastRow();
  var namaPelajar = (data.studentName || "").toString().trim();

  var barisSama = 0;
  var urlLama = [];
  var barisBuang = [];

  if (lastRow > 1) {
    var records = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < records.length; i++) {
      var rowNama = records[i][1] ? records[i][1].toString().trim() : "";
      if (rowNama !== namaPelajar) continue;
      if (!barisSama) {
        barisSama = i + 2;          // guna baris pertama pelajar ini sebagai rekod terkini
        urlLama.push(records[i][4]); // fail lama pada baris ini juga perlu dipadam
      } else {
        urlLama.push(records[i][4]);
        barisBuang.push(i + 2);      // baris sejarah lain -> buang
      }
    }
  }

  // 3) Kemas kini rekod supaya menunjuk kepada fail TERKINI.
  if (barisSama) {
    sheet.getRange(barisSama, 1, 1, 6).setValues([[tarikh, namaPelajar, data.certName, data.guruEmail, fileUrl, dikemaskiniOleh]]);
  } else {
    sheet.appendRow([tarikh, namaPelajar, data.certName, data.guruEmail, fileUrl, dikemaskiniOleh]);
  }

  // 4) Barulah padam baris sejarah + fail PDF lama di Drive.
  barisBuang.sort(function (a, b) { return b - a; });
  for (var d = 0; d < barisBuang.length; d++) {
    try { sheet.deleteRow(barisBuang[d]); } catch (e) {}
  }
  var dipadam = 0;
  for (var u = 0; u < urlLama.length; u++) {
    if (padamFailDrive(urlLama[u], newId)) dipadam++;
  }

  CacheService.getScriptCache().remove("certs_" + namaPelajar);
  return {
    status: "success",
    message: "Fail terkini disimpan" + (dipadam ? " (" + dipadam + " fail lama dipadam)." : "."),
    url: fileUrl,
    deleted: dipadam
  };
}


function dapatkanSijil(studentName) {
  studentName = (studentName || "").toString().trim();
  var cache = CacheService.getScriptCache();
  var key = "certs_" + studentName;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = getOrCreateSheet(ss, SHEET_SIJIL);
  pastikanLajurSijil(sheet);
  var lastRow = sheet.getLastRow();
  var certs = [];
  if (lastRow > 1) {
    // Cari terus pada lajur nama; elak memindahkan semua rekod sijil ke memori.
    var matches = sheet.getRange(2, 2, lastRow - 1, 1)
      .createTextFinder(studentName)
      .matchEntireCell(true)
      .matchCase(true)
      .findAll();
    for (var i = matches.length - 1; i >= 0; i--) {
      var row = sheet.getRange(matches[i].getRow(), 1, 1, 6).getValues()[0];
      certs.push({ date: row[0], certName: row[2], url: row[4], guruEmail: row[3], updatedBy: row[5] || row[3] });
    }
  }
  var res = { status: "success", data: certs };
  try { cache.put(key, JSON.stringify(res), CACHE_SEC); } catch (e) {}
  return res;
}

function tarikPdfUntukDiedit(url) {
  try {
    if (!url) return { status: "error", message: "Pautan fail kosong." };
    var idMatch = String(url).match(/[-\w]{25,}/);
    if (!idMatch) return { status: "error", message: "ID fail tidak sah." };

    var file = DriveApp.getFileById(idMatch[0]);
    var mime = file.getMimeType();
    var blob;

    // Google Docs/Slides bukan fail binari - eksport dulu jadi PDF
    if (mime && mime.indexOf("application/vnd.google-apps") === 0) {
      if (mime === "application/vnd.google-apps.folder") {
        return { status: "error", message: "Pautan ini adalah folder, bukan fail sijil." };
      }
      blob = file.getAs("application/pdf");
      mime = "application/pdf";
    } else {
      blob = file.getBlob();
    }

    var bytes = blob.getBytes();
    if (!bytes || bytes.length === 0) {
      return { status: "error", message: "Fail kosong (0 bait) di Drive." };
    }
    // Had selamat Apps Script (~50MB respons); elak respons terpotong yang buat atob gagal
    if (bytes.length > 20 * 1024 * 1024) {
      return { status: "error", message: "Fail terlalu besar untuk diedit (" + Math.round(bytes.length / 1048576) + " MB). Had 20 MB." };
    }

    return {
      status: "success",
      base64Data: Utilities.base64Encode(bytes),
      mimeType: mime,
      fileName: file.getName(),
      size: bytes.length
    };
  } catch (err) {
    return { status: "error", message: "Gagal akses fail: " + err.toString() };
  }
}

/* ============================================================
 * PELAJAR (dioptimumkan untuk ribuan rekod)
 * ============================================================ */
/**
 * data: { search, page (1-based), limit, sort ('az'|'za') }
 * Pulangkan hanya sekeping (page) data -> pantas walaupun 10,000 pelajar.
 */
function dapatkanPelajar(data) {
  data = data || {};
  var search = (data.search || "").toString().trim().toLowerCase();
  var page = parseInt(data.page, 10) || 1;
  var limit = parseInt(data.limit, 10) || 100;
  var sort = data.sort === "za" ? "za" : "az";

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var all = bacaSemuaPelajarRingkas(sheet);

  if (search) {
    all = all.filter(function (s) {
      return s.name.toLowerCase().indexOf(search) > -1 ||
        s.ic.toString().toLowerCase().indexOf(search) > -1 ||
        s.kelas.toString().toLowerCase().indexOf(search) > -1 ||
        (s.sekolah || "").toString().toLowerCase().indexOf(search) > -1;
    });
  }

  all.sort(function (a, b) {
    return sort === "za" ? b.name.localeCompare(a.name) : a.name.localeCompare(b.name);
  });

  var total = all.length;
  var start = (page - 1) * limit;
  var slice = all.slice(start, start + limit);

  return {
    status: "success",
    data: slice,
    total: total,
    page: page,
    limit: limit,
    hasMore: start + slice.length < total
  };
}

function tambahPelajar(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var nameTrimmed = data.studentName.trim();
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (names[i][0] && names[i][0].toString().trim().toLowerCase() === nameTrimmed.toLowerCase()) {
        return { status: "error", message: "Nama pelajar ini sudah direkodkan di dalam sistem!" };
      }
    }
  }

  var gambarId = simpanGambarProfil(data.gambarBase64, nameTrimmed);
  var tarikh = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm");
  pastikanLajurSekolah(sheet);
  sheet.appendRow([nameTrimmed, data.kelas || "-", data.ic || "-", data.jantina || "-", data.kategori || "-", gambarId, data.guruEmail, tarikh, data.sekolah || "-"]);
  return { status: "success", message: "Pelajar berjaya didaftarkan.", gambarId: gambarId };
}

function kemaskiniMaklumatPelajar(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var oldName = data.oldName.trim();
  var newName = data.studentName.trim();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "error", message: "Pelajar tidak dijumpai." };

  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  if (oldName.toLowerCase() !== newName.toLowerCase()) {
    for (var j = 0; j < names.length; j++) {
      if (names[j][0] && names[j][0].toString().trim().toLowerCase() === newName.toLowerCase()) {
        return { status: "error", message: "Gagal: Nama baru ini sudah wujud dalam sistem." };
      }
    }
  }

  for (var i = 0; i < names.length; i++) {
    if (names[i][0] && names[i][0].toString().trim() === oldName) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, 1, 1, 5).setValues([[newName, data.kelas || "-", data.ic || "-", data.jantina || "-", data.kategori || "-"]]);
      pastikanLajurSekolah(sheet);
      sheet.getRange(rowNum, 9).setValue(data.sekolah || "-");

      var gambarId = "";
      if (data.gambarBase64) {
        gambarId = simpanGambarProfil(data.gambarBase64, newName);
        if (gambarId) sheet.getRange(rowNum, 6).setValue(gambarId);
      } else {
        var existing = sheet.getRange(rowNum, 6).getValue().toString();
        // Migrasi rekod lama base64 -> fail Drive
        if (existing.indexOf("data:") === 0) {
          gambarId = simpanGambarProfil(existing, newName);
          sheet.getRange(rowNum, 6).setValue(gambarId);
        } else {
          gambarId = existing;
        }
      }

      if (oldName !== newName) {
        var sheetSijil = getOrCreateSheet(ss, SHEET_SIJIL);
        var lastSijil = sheetSijil.getLastRow();
        if (lastSijil > 1) {
          var certNames = sheetSijil.getRange(2, 2, lastSijil - 1, 1).getValues();
          for (var k = 0; k < certNames.length; k++) {
            if (certNames[k][0] && certNames[k][0].toString().trim() === oldName) {
              sheetSijil.getRange(k + 2, 2).setValue(newName);
            }
          }
        }
        CacheService.getScriptCache().remove("certs_" + oldName);
      }
      CacheService.getScriptCache().remove("certs_" + newName);
      return { status: "success", message: "Profil pelajar berjaya dikemas kini!", gambarId: gambarId };
    }
  }
  return { status: "error", message: "Pelajar tidak dijumpai dalam pangkalan data." };
}

/**
 * Semakan borang oleh guru.
 * data: { studentName, semakan: 'lengkap'|'belum'|'', catatan, guruNama, guruEmail }
 */
function tetapSemakan(data) {
  data = data || {};
  var nama = (data.studentName || "").toString().trim();
  if (!nama) return { status: "error", message: "Nama pelajar diperlukan." };
  var status = (data.semakan || "").toString().toLowerCase();
  if (status !== "lengkap" && status !== "belum") status = "";
  var catatan = status === "belum" ? (data.catatan || "").toString().trim() : "";

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  pastikanLajurSemakan(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "error", message: "Tiada rekod pelajar." };

  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] && names[i][0].toString().trim() === nama) {
      var oleh = status ? ((data.guruNama || data.guruEmail || "").toString()) : "";
      var tarikh = status ? Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd-MM-yyyy HH:mm") : "";
      sheet.getRange(i + 2, 10, 1, 4).setValues([[status, catatan, oleh, tarikh]]);
      return {
        status: "success",
        message: status === "lengkap" ? "Ditanda LENGKAP." : (status === "belum" ? "Ditanda BELUM LENGKAP." : "Tanda semakan dibuang."),
        semakan: status, catatan: catatan, semakOleh: oleh, semakTarikh: tarikh
      };
    }
  }
  return { status: "error", message: "Pelajar tidak dijumpai." };
}

function buangPelajar(studentName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var nama = (studentName || "").toString().trim();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: "error", message: "Tiada rekod." };

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var jumpa = false;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] && rows[i][0].toString().trim() === nama) {
      // padam gambar profil di Drive (jika fail, bukan base64 lama)
      var gid = rows[i][5] ? rows[i][5].toString() : "";
      if (gid && gid.indexOf("data:") !== 0) {
        try { DriveApp.getFileById(gid).setTrashed(true); } catch (e) {}
      }
      sheet.deleteRow(i + 2);
      jumpa = true;
    }
  }
  if (!jumpa) return { status: "error", message: "Pelajar tidak dijumpai." };

  // Padam semua sijil pelajar ini (rekod sheet + fail PDF di Drive)
  var sijilDipadam = 0;
  var sheetSijil = getOrCreateSheet(ss, SHEET_SIJIL);
  var lastSijil = sheetSijil.getLastRow();
  if (lastSijil > 1) {
    var certs = sheetSijil.getRange(2, 1, lastSijil - 1, 5).getValues();
    for (var k = certs.length - 1; k >= 0; k--) {
      var rowNama = certs[k][1] ? certs[k][1].toString().trim() : "";
      if (rowNama !== nama) continue;
      padamFailDrive(certs[k][4], null);
      try { sheetSijil.deleteRow(k + 2); } catch (e) {}
      sijilDipadam++;
    }
  }

  CacheService.getScriptCache().remove("certs_" + nama);
  return {
    status: "success",
    message: "Rekod pelajar dibuang" + (sijilDipadam ? " bersama " + sijilDipadam + " sijil." : "."),
    certsDeleted: sijilDipadam
  };
}


/**
 * JALANKAN SEKALI SAHAJA (manual dari editor Apps Script)
 * Memindahkan gambar Base64 lama dalam Sheet ke fail Drive supaya sistem laju.
 */
function migrasiGambarBase64KeDrive() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_PELAJAR);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  for (var r = 2; r <= lastRow; r++) {
    var val = sheet.getRange(r, 6).getValue().toString();
    if (val.indexOf("data:") === 0) {
      var nama = sheet.getRange(r, 1).getValue().toString();
      var id = simpanGambarProfil(val, nama);
      sheet.getRange(r, 6).setValue(id);
    }
  }
}
