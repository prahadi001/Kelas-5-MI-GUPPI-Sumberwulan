// ============================================================
//  PORTAL UJIAN DIGITAL — MI GUPPI SUMBERWULAN
//  code.gs  |  Versi 2.1
// ============================================================

var CONFIG = {
  folderID   : "1I-rB4XC9DsTPEe4-h7HDfmZ-AOu7fUZd",
  sheetSoal  : "Bank Soal",
  sheetSiswa : "Siswa",
  sheetNilai : "NilaiSiswa",
  sheetConfig: "Konfigurasi",
  sheetMateri: "MateriBelajar"
};

// ==================== ENTRY POINT ====================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Ujian Digital — MI GUPPI Sumberwulan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== HELPER ====================
function _getOrCreateSheet(nama, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(nama);
  if (!sh) {
    sh = ss.insertSheet(nama);
    if (headers && headers.length > 0) sh.appendRow(headers);
  }
  return sh;
}

function _sanitize(str) {
  if (typeof str !== 'string') return String(str || '');
  return str.trim().replace(/<[^>]*>/g, '');
}

// ==================== KONFIGURASI ====================
function _initKonfigurasi() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.sheetConfig);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.sheetConfig);
    sh.appendRow(["Kunci", "Nilai"]);
    sh.appendRow(["password_guru",      "guru123"]);
    sh.appendRow(["durasi_sesi_menit",  "60"]);
    sh.appendRow(["durasi_ujian_menit", "90"]);
    sh.appendRow(["nilai_kkm",          "70"]);
    sh.appendRow(["nama_sekolah",       "MI GUPPI Sumberwulan"]);
  }
  return sh;
}

function getKonfigurasi() {
  try {
    var sh   = _initKonfigurasi();
    var data = sh.getDataRange().getValues();
    var cfg  = {};
    for (var i = 1; i < data.length; i++) {
      cfg[data[i][0]] = data[i][1];
    }
    delete cfg['password_guru'];
    return cfg;
  } catch(e) {
    return { durasi_sesi_menit:60, durasi_ujian_menit:90, nilai_kkm:70, nama_sekolah:"MI GUPPI Sumberwulan" };
  }
}

function _getKKM() {
  try {
    var sh   = _initKonfigurasi();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "nilai_kkm") return parseInt(data[i][1]) || 70;
    }
    return 70;
  } catch(e) { return 70; }
}

function _getGuruPassword() {
  try {
    var sh   = _initKonfigurasi();
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "password_guru") return data[i][1].toString();
    }
    return "guru123";
  } catch(e) { return "guru123"; }
}

// ==================== LOGIN ====================
function checkLogin(username, password) {
  if (!username || !password) return { success: false, message: "Username dan password wajib diisi!" };
  var userClean = _sanitize(username).toLowerCase();
  var passClean = _sanitize(password);

  // Cek guru
  if (userClean === "guru" && passClean === _getGuruPassword()) {
    return { success: true, role: "guru", namaResmi: "Bapak/Ibu Guru" };
  }

  // Cek siswa
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.sheetSiswa);
    if (!sh) return { success: false, message: "Sheet Siswa belum dibuat. Hubungi guru." };
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var nama  = data[i][0] ? data[i][0].toString().trim() : "";
      var pass  = data[i][1] ? data[i][1].toString().trim() : "";
      if (nama.toLowerCase() === userClean && pass === passClean) {
        return { success: true, role: "siswa", namaResmi: nama };
      }
    }
    return { success: false, message: "Username atau password salah." };
  } catch(err) {
    return { success: false, message: "Error: " + err.message };
  }
}

// ==================== MANAJEMEN SISWA ====================
function getDaftarSiswa() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.sheetSiswa);
    if (!sh) return [];
    var data   = sh.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0] === "") continue;
      result.push({ baris: i+1, nama: data[i][0], pass: data[i][1] });
    }
    return result;
  } catch(e) { return []; }
}

function tambahSiswa(nama, password) {
  try {
    nama     = _sanitize(nama);
    password = _sanitize(password);
    if (!nama || !password) return { success: false, message: "Nama dan password wajib diisi!" };
    var sh   = _getOrCreateSheet(CONFIG.sheetSiswa, ["Nama Siswa","Password","Dibuat"]);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === nama.toLowerCase()) {
        return { success: false, message: "Nama siswa sudah terdaftar!" };
      }
    }
    sh.appendRow([nama, password, Utilities.formatDate(new Date(),"GMT+7","yyyy-MM-dd HH:mm:ss")]);
    return { success: true, message: "Siswa berhasil ditambahkan." };
  } catch(e) { return { success: false, message: e.message }; }
}

function hapusSiswa(baris) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.sheetSiswa);
    if (!sh) return { success: false, message: "Sheet tidak ditemukan." };
    sh.deleteRow(parseInt(baris));
    return { success: true, message: "Data siswa berhasil dihapus." };
  } catch(e) { return { success: false, message: e.message }; }
}

function ubahPasswordSiswa(baris, passwordBaru) {
  try {
    passwordBaru = _sanitize(passwordBaru);
    if (!passwordBaru) return { success: false, message: "Password baru tidak boleh kosong!" };
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.sheetSiswa);
    if (!sh) return { success: false, message: "Sheet tidak ditemukan." };
    sh.getRange(parseInt(baris), 2).setValue(passwordBaru);
    return { success: true, message: "Password berhasil diubah." };
  } catch(e) { return { success: false, message: e.message }; }
}

// ==================== BANK SOAL ====================
function uploadFiles(formObject) {
  try {
    var folder = DriveApp.getFolderById(CONFIG.folderID);
    var sheet  = _getOrCreateSheet(CONFIG.sheetSoal, [
      "Mata Pelajaran","Semester","BAB","Materi","Jenis Soal",
      "Pertanyaan","Pilihan A","Pilihan B","Pilihan C","Pilihan D",
      "Kunci","Bobot","Nama File","URL File","Timestamp","Link Youtube"
    ]);

    var fileUrl  = "Tidak ada lampiran";
    var fileName = "";
    if (formObject.myFile && formObject.myFile.getName && formObject.myFile.getName() !== "") {
      var blob = formObject.myFile;
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl  = file.getUrl();
      fileName = file.getName();
    }

    var jenis = _sanitize(formObject.jenis_soal || "Pilihan Ganda");
    var kunci = "";
    if (jenis === "Pilihan Ganda") {
      kunci = _sanitize(formObject.kunci_pg || "").toUpperCase();
    } else if (jenis === "Isian") {
      kunci = _sanitize(formObject.kunci_isian || "");
    } else {
      kunci = _sanitize(formObject.kunci_uraian || "");
    }

    sheet.appendRow([
      _sanitize(formObject.mata_pelajaran),
      _sanitize(formObject.semester),
      _sanitize(formObject.bab_soal),
      _sanitize(formObject.materi_soal || ""),
      jenis,
      _sanitize(formObject.teks_pertanyaan),
      _sanitize(formObject.pilihan_a || ""),
      _sanitize(formObject.pilihan_b || ""),
      _sanitize(formObject.pilihan_c || ""),
      _sanitize(formObject.pilihan_d || ""),
      kunci,
      parseFloat(formObject.bobot_nilai) || 10,
      fileName,
      fileUrl,
      Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
      _sanitize(formObject.link_youtube || "")
    ]);
    return { success: true, message: "Soal berhasil disimpan." };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

// FUNGSI UTAMA — ambil semua soal dari sheet "Bank Soal"
function getBankSoal() {
  try {
    var ss    = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CONFIG.sheetSoal);

    // Sheet belum ada → kembalikan array kosong (bukan error)
    if (!sheet) {
      Logger.log("getBankSoal: Sheet 'Bank Soal' tidak ditemukan.");
      return [];
    }

    var lastRow = sheet.getLastRow();
    Logger.log("getBankSoal: lastRow = " + lastRow);

    // Hanya header atau kosong
    if (lastRow <= 1) return [];

    // Ambil semua data sekaligus (lebih efisien)
    var data   = sheet.getRange(1, 1, lastRow, 16).getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      // Lewati baris kosong (kolom pertama = Mata Pelajaran harus ada)
      var mapel = data[i][0] ? data[i][0].toString().trim() : "";
      if (mapel === "") continue;

      var rawUrl   = data[i][13] ? data[i][13].toString() : "";
      var fileName = data[i][12] ? data[i][12].toString() : "";

      result.push({
        id         : i,
        mapel      : data[i][0]  ? data[i][0].toString().trim()  : "",
        semester   : data[i][1]  ? data[i][1].toString().trim()  : "",
        bab        : data[i][2]  ? data[i][2].toString().trim()  : "",
        materi     : data[i][3]  ? data[i][3].toString().trim()  : "",
        jenis      : data[i][4]  ? data[i][4].toString().trim()  : "Pilihan Ganda",
        pertanyaan : data[i][5]  ? data[i][5].toString().trim()  : "",
        pilihan_a  : data[i][6]  ? data[i][6].toString().trim()  : "",
        pilihan_b  : data[i][7]  ? data[i][7].toString().trim()  : "",
        pilihan_c  : data[i][8]  ? data[i][8].toString().trim()  : "",
        pilihan_d  : data[i][9]  ? data[i][9].toString().trim()  : "",
        kunci      : data[i][10] ? data[i][10].toString().trim() : "",
        bobot      : data[i][11] ? parseFloat(data[i][11]) || 10 : 10,
        fileName   : fileName,
        fileUrl    : _konversiUrlDrive(rawUrl, fileName),
        youtube    : data[i][15] ? data[i][15].toString().trim() : ""
      });
    }

    Logger.log("getBankSoal: total soal = " + result.length);
    return result;

  } catch(e) {
    Logger.log("getBankSoal ERROR: " + e.toString());
    // Kembalikan error sebagai objek agar bisa ditampilkan di frontend
    return { error: true, message: e.toString() };
  }
}

function hapusSoal(idBaris) {
  try {
    var ss    = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CONFIG.sheetSoal);
    if (!sheet) return { success: false, message: "Sheet Bank Soal tidak ditemukan." };
    // idBaris = index di array result (0-based dari data row 1), baris sheet = idBaris + 2
    sheet.deleteRow(parseInt(idBaris) + 2);
    return { success: true, message: "Soal berhasil dihapus." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
//  KELAS BELAJAR (Materi Pembelajaran per Mata Pelajaran)
// ============================================================
function simpanMateri(formObject) {
  try {
    var folder = DriveApp.getFolderById(CONFIG.folderID);
    var sheet  = _getOrCreateSheet(CONFIG.sheetMateri, [
      "Mata Pelajaran","Judul Materi","Isi / Deskripsi","Jenis File","Nama File","URL File","Timestamp","Link Youtube"
    ]);

    var fileUrl  = "Tidak ada lampiran";
    var fileName = "";
    var jenisFile = "-";
    if (formObject.myFile && formObject.myFile.getName && formObject.myFile.getName() !== "") {
      var blob = formObject.myFile;
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl   = file.getUrl();
      fileName  = file.getName();
      if (/\.(jpg|jpeg|png|gif|webp)$/i.test(fileName)) jenisFile = "Gambar";
      else if (/\.(mp4|webm|mov)$/i.test(fileName)) jenisFile = "Video";
      else jenisFile = "Dokumen";
    }

    if (!_sanitize(formObject.mata_pelajaran)) {
      return { success: false, message: "Mata Pelajaran wajib dipilih." };
    }
    if (!_sanitize(formObject.judul_materi)) {
      return { success: false, message: "Judul materi wajib diisi." };
    }

    sheet.appendRow([
      _sanitize(formObject.mata_pelajaran),
      _sanitize(formObject.judul_materi),
      _sanitize(formObject.isi_materi || ""),
      jenisFile,
      fileName,
      fileUrl,
      Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss"),
      _sanitize(formObject.link_youtube || "")
    ]);
    return { success: true, message: "Materi berhasil disimpan." };
  } catch(error) {
    return { success: false, message: error.toString() };
  }
}

function getSemuaMateri() {
  try {
    var ss    = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CONFIG.sheetMateri);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data   = sheet.getRange(1, 1, lastRow, 8).getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var mapel = data[i][0] ? data[i][0].toString().trim() : "";
      if (mapel === "") continue;

      var rawUrl   = data[i][5] ? data[i][5].toString() : "";
      var fileName = data[i][4] ? data[i][4].toString() : "";

      result.push({
        id       : i,
        mapel    : mapel,
        judul    : data[i][1] ? data[i][1].toString().trim() : "",
        isi      : data[i][2] ? data[i][2].toString().trim() : "",
        jenisFile: data[i][3] ? data[i][3].toString().trim() : "-",
        fileName : fileName,
        fileUrl  : _konversiUrlDrive(rawUrl, fileName),
        timestamp: data[i][6] ? data[i][6].toString() : "",
        youtube  : data[i][7] ? data[i][7].toString().trim() : ""
      });
    }

    return result.reverse();
  } catch(e) {
    Logger.log("getSemuaMateri ERROR: " + e.toString());
    return { error: true, message: e.toString() };
  }
}

function hapusMateri(idBaris) {
  try {
    var ss    = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(CONFIG.sheetMateri);
    if (!sheet) return { success: false, message: "Sheet Materi Belajar tidak ditemukan." };
    sheet.deleteRow(parseInt(idBaris) + 2);
    return { success: true, message: "Materi berhasil dihapus." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function _konversiUrlDrive(rawUrl, fileName) {
  if (!rawUrl || rawUrl === "" || rawUrl === "Tidak ada lampiran") return "";
  if (rawUrl.indexOf('drive.google.com') === -1) return rawUrl;
  try {
    var fileId = "";
    if (rawUrl.indexOf('/d/') !== -1) {
      fileId = rawUrl.split('/d/')[1].split('/')[0];
    } else if (rawUrl.indexOf('id=') !== -1) {
      fileId = rawUrl.split('id=')[1].split('&')[0];
    }
    if (!fileId) return rawUrl;
    var ext = (fileName || "").toLowerCase();
    if (/\.(jpg|jpeg|png|gif|webp)$/.test(ext)) {
      return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
    } else if (/\.(mp4|webm|mov)$/.test(ext)) {
      return "https://drive.google.com/file/d/" + fileId + "/preview";
    } else {
      return "https://drive.google.com/file/d/" + fileId + "/view";
    }
  } catch(e) { return rawUrl; }
}

// ==================== NILAI ====================
function simpanNilaiSiswaToServer(dataNilai) {
  try {
    var sheet = _getOrCreateSheet(CONFIG.sheetNilai, [
      "Timestamp","Nama Siswa","Mata Pelajaran","BAB","Materi",
      "Semester","Jumlah Soal","Benar","Salah","Kosong","Nilai Akhir","Lulus/Tidak","Jawaban Uraian"
    ]);
    var kkm   = _getKKM();
    var lulus = (parseFloat(dataNilai.nilaiAkhir) >= kkm) ? "LULUS" : "TIDAK LULUS";
    sheet.appendRow([
      new Date(),
      _sanitize(dataNilai.namaSiswa || ""),
      _sanitize(dataNilai.mapel     || ""),
      _sanitize(dataNilai.bab       || "-"),
      _sanitize(dataNilai.materi    || "-"),
      _sanitize(dataNilai.semester  || ""),
      parseInt(dataNilai.jumlahSoal || 0),
      parseInt(dataNilai.benar      || 0),
      parseInt(dataNilai.salah      || 0),
      parseInt(dataNilai.kosong     || 0),
      parseFloat(dataNilai.nilaiAkhir || 0),
      lulus,
      _sanitize(dataNilai.jawabanUraian || "")
    ]);
    return { success: true, message: "Nilai berhasil disimpan!", kkm: kkm, lulus: lulus };
  } catch(error) {
    return { success: false, message: "Gagal: " + error.toString() };
  }
}

function getRekapNilai() {
  try {
    var ss    = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("NilaiSiswa");
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var kkm  = _getKKM();
    // Baca semua data (baris 2 sampai akhir = baris 1 ke lastRow-1 di array)
    var data = sheet.getDataRange().getValues();

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var b = data[i];

      // Kolom A-L: Timestamp,NamaSiswa,MapPel,BAB,Materi,Semester,JmlSoal,Benar,Salah,Kosong,Nilai,Lulus
      var nama = String(b[1] || "").trim();
      if (!nama) continue;

      // Nilai bisa 0, parse dengan aman
      var nilaiNum = 0;
      if (b[10] !== "" && b[10] !== null && b[10] !== undefined) {
        nilaiNum = parseFloat(b[10]);
        if (isNaN(nilaiNum)) nilaiNum = 0;
      }

      var lulus = String(b[11] || "").trim();
      if (!lulus) lulus = nilaiNum >= kkm ? "LULUS" : "TIDAK LULUS";

      result.push({
        "Timestamp"     : b[0] || "",
        "Nama Siswa"    : nama,
        "Mata Pelajaran": String(b[2] || "").trim(),
        "BAB"           : String(b[3] || "-").trim() || "-",
        "Materi"        : String(b[4] || "-").trim() || "-",
        "Semester"      : String(b[5] || "").trim(),
        "Jumlah Soal"   : Number(b[6]) || 0,
        "Benar"         : Number(b[7]) || 0,
        "Salah"         : Number(b[8]) || 0,
        "Kosong"        : Number(b[9]) || 0,
        "Nilai Akhir"   : nilaiNum,
        "Lulus/Tidak"   : lulus,
        "Jawaban Uraian": String(b[12] || "-").trim() || "-"
      });
    }

    return result.reverse();
  } catch(e) {
    Logger.log("getRekapNilai ERROR: " + e);
    return [];
  }
}

// Diagnostik — cek semua sheet yang ada di spreadsheet
function getDiagnostik() {
  try {
    var ss      = SpreadsheetApp.getActive();
    var sheets  = ss.getSheets();
    var info    = [];
    for (var i = 0; i < sheets.length; i++) {
      info.push({
        nama    : sheets[i].getName(),
        baris   : sheets[i].getLastRow(),
        kolom   : sheets[i].getLastColumn()
      });
    }
    return {
      namaFile    : ss.getName(),
      jumlahSheet : sheets.length,
      sheets      : info,
      configNilai : CONFIG.sheetNilai,
      configSoal  : CONFIG.sheetSoal,
      configSiswa : CONFIG.sheetSiswa
    };
  } catch(e) {
    return { error: e.toString() };
  }
}

// Input nilai manual oleh guru (tanpa perlu siswa mengerjakan soal)
function tambahNilaiManual(d) {
  try {
    var sheet = _getOrCreateSheet(CONFIG.sheetNilai, [
      "Timestamp","Nama Siswa","Mata Pelajaran","BAB","Materi",
      "Semester","Jumlah Soal","Benar","Salah","Kosong","Nilai Akhir","Lulus/Tidak","Jawaban Uraian"
    ]);
    var kkm   = _getKKM();
    var nilai = parseFloat(d.nilaiAkhir) || 0;
    var lulus = nilai >= kkm ? "LULUS" : "TIDAK LULUS";
    sheet.appendRow([
      new Date(),
      _sanitize(d.namaSiswa   || ""),
      _sanitize(d.mapel       || ""),
      _sanitize(d.bab         || "-"),
      _sanitize(d.materi      || "-"),
      _sanitize(d.semester    || ""),
      parseInt(d.jumlahSoal   || 0),
      parseInt(d.benar        || 0),
      parseInt(d.salah        || 0),
      parseInt(d.kosong       || 0),
      nilai,
      lulus,
      "-"
    ]);
    return { success: true, message: "Nilai berhasil ditambahkan. Status: " + lulus };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getRiwayatNilaiSiswa(namaSiswa) {
  try {
    // Pakai Raw agar konsisten dengan getRekapNilaiRaw
    var semua  = getRekapNilaiRaw();
    if (!Array.isArray(semua)) return [];
    var bersih = _sanitize(namaSiswa || "").toLowerCase();
    if (!bersih) return semua;
    return semua.filter(function(r) {
      // getRekapNilaiRaw pakai field "nama"
      var n = (r.nama || r["Nama Siswa"] || "").toString().toLowerCase();
      return n === bersih;
    });
  } catch(e) {
    Logger.log("getRiwayatNilaiSiswa ERROR: " + e.toString());
    return [];
  }
}

// ==================== STATISTIK ====================
function getStatistikDashboard() {
  try {
    var ss = SpreadsheetApp.getActive();
    var shSoal  = ss.getSheetByName(CONFIG.sheetSoal);
    var shSiswa = ss.getSheetByName(CONFIG.sheetSiswa);
    var shNilai = ss.getSheetByName(CONFIG.sheetNilai);
    var jmlSoal  = shSoal  ? Math.max(0, shSoal.getLastRow()  - 1) : 0;
    var jmlSiswa = shSiswa ? Math.max(0, shSiswa.getLastRow() - 1) : 0;
    var jmlUjian = shNilai ? Math.max(0, shNilai.getLastRow() - 1) : 0;
    var rataRata = 0;
    if (jmlUjian > 0 && shNilai) {
      var dataNilai = shNilai.getDataRange().getValues();
      var total = 0;
      for (var i = 1; i < dataNilai.length; i++) total += parseFloat(dataNilai[i][9]) || 0;
      rataRata = Math.round((total / jmlUjian) * 10) / 10;
    }
    return { jmlSoal:jmlSoal, jmlSiswa:jmlSiswa, jmlUjian:jmlUjian, rataRata:rataRata, kkm:_getKKM() };
  } catch(e) {
    return { jmlSoal:0, jmlSiswa:0, jmlUjian:0, rataRata:0, kkm:70 };
  }
}

// ==================== FUNGSI TEST (Jalankan dari Apps Script Editor) ====================
// Cara pakai: Buka Apps Script > pilih fungsi "testRekapNilai" > klik ▶ Run
// Lihat hasilnya di menu View > Logs (Ctrl+Enter)
function testRekapNilai() {
  var hasil = getRekapNilai();
  Logger.log("=== TEST getRekapNilai ===");
  Logger.log("Tipe hasil: " + typeof hasil);
  Logger.log("Array? " + Array.isArray(hasil));
  if (Array.isArray(hasil)) {
    Logger.log("Jumlah baris: " + hasil.length);
    if (hasil.length > 0) {
      Logger.log("Baris pertama: " + JSON.stringify(hasil[0]));
    }
  } else {
    Logger.log("Isi: " + JSON.stringify(hasil));
  }

  // Cek sheet langsung
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("NilaiSiswa");
  if (sheet) {
    Logger.log("Sheet NilaiSiswa: lastRow=" + sheet.getLastRow() + " lastCol=" + sheet.getLastColumn());
    var data = sheet.getDataRange().getValues();
    Logger.log("Header: " + JSON.stringify(data[0]));
    Logger.log("Baris-2: " + JSON.stringify(data[1]));
  }
}

// ── FUNGSI PALING SEDERHANA — fallback jika getRekapNilai bermasalah ──
function getRekapNilaiRaw() {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("NilaiSiswa");
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data   = sheet.getDataRange().getValues();
  var kkm    = _getKKM();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    // Lewati baris benar-benar kosong
    if (!r[1] || r[1].toString().trim() === "") continue;

    var nilai = 0;
    if (r[10] !== "" && r[10] !== null && r[10] !== undefined) {
      nilai = parseFloat(r[10]);
      if (isNaN(nilai)) nilai = 0;
    }

    var status = r[11] ? r[11].toString().trim() : "";
    if (!status) status = nilai >= kkm ? "LULUS" : "TIDAK LULUS";

    result.push({
      ts     : r[0] ? r[0].toString() : "",
      nama   : r[1].toString().trim(),
      mapel  : r[2] ? r[2].toString().trim() : "",
      bab    : r[3] ? r[3].toString().trim() : "-",
      materi : r[4] ? r[4].toString().trim() : "-",
      sem    : r[5] ? r[5].toString().trim() : "",
      jml    : Number(r[6]) || 0,
      benar  : Number(r[7]) || 0,
      salah  : Number(r[8]) || 0,
      kosong : Number(r[9]) || 0,
      nilai  : nilai,
      status : status,
      jawabanUraian: r[12] ? r[12].toString().trim() : "-"
    });
  }

  result.reverse();
  return result;
}

// ============================================================
//  RESET SHEET LAMA -> BUAT ULANG DENGAN HEADER BARU (TANPA KELAS)
// ============================================================
// PERINGATAN: Fungsi ini MENGHAPUS seluruh isi sheet "Bank Soal",
// "Siswa", dan "NilaiSiswa" beserta datanya, lalu membuatnya ulang
// dari nol dengan header baru (kolom "Kelas" sudah tidak ada).
//
// Cara pakai:
// 1. Buka project Apps Script (Extensions > Apps Script).
// 2. Pilih fungsi "resetDataKelas5" di dropdown toolbar.
// 3. Klik ▶ Run, lalu izinkan otorisasi jika diminta.
// 4. Lihat hasilnya di menu View > Logs (Ctrl+Enter).
// 5. Cukup dijalankan SEKALI SAJA. Jangan dijalankan berulang
//    kali jika sudah ada data baru yang ingin dipertahankan.
function resetDataKelas5() {
  var ss = SpreadsheetApp.getActive();

  var daftarSheet = [
    {
      nama: CONFIG.sheetSoal,
      headers: [
        "Mata Pelajaran","Semester","BAB","Materi","Jenis Soal",
        "Pertanyaan","Pilihan A","Pilihan B","Pilihan C","Pilihan D",
        "Kunci","Bobot","Nama File","URL File","Timestamp","Link Youtube"
      ]
    },
    {
      nama: CONFIG.sheetSiswa,
      headers: ["Nama Siswa","Password","Dibuat"]
    },
    {
      nama: CONFIG.sheetNilai,
      headers: [
        "Timestamp","Nama Siswa","Mata Pelajaran","BAB","Materi",
        "Semester","Jumlah Soal","Benar","Salah","Kosong","Nilai Akhir","Lulus/Tidak","Jawaban Uraian"
      ]
    }
  ];

  var log = [];
  daftarSheet.forEach(function(item) {
    var shLama = ss.getSheetByName(item.nama);
    if (shLama) {
      ss.deleteSheet(shLama);
      log.push('Sheet "' + item.nama + '" lama dihapus.');
    }
    var shBaru = ss.insertSheet(item.nama);
    shBaru.appendRow(item.headers);
    shBaru.getRange(1, 1, 1, item.headers.length).setFontWeight("bold");
    log.push('Sheet "' + item.nama + '" dibuat ulang dengan header: ' + item.headers.join(", "));
  });

  Logger.log(log.join("\n"));
  return { success: true, message: "Reset selesai. Semua sheet sudah dibuat ulang dengan header baru (tanpa kolom Kelas)." };
}