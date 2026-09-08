/**
 * @file CustomMenu.gs
 * @description Menambahkan menu kustom Dapodik di Google Sheets UI.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🎓 Dapodik Kemendikdasmen")
    .addItem("📥 Tarik Semua Siswa (Peserta Didik)", "menuSyncSiswa")
    .addItem("📥 Tarik Semua Guru & Tendik (GTK)", "menuSyncGtk")
    .addItem("📥 Tarik Profil Sekolah", "menuSyncSekolah")
    .addItem("📥 Tarik Rombongan Belajar (Rombel)", "menuSyncRombel")
    .addItem("📥 Tarik Prasarana & Bangunan", "menuSyncPrasarana")
    .addSeparator()
    .addItem("⚙️ Setup Sheet Konfigurasi Token", "menuSetupConfigSheet")
    .addToUi();
}

function menuSetupConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Config");
  if (!sheet) {
    sheet = ss.insertSheet("Config");
  } else {
    sheet.clear();
  }

  var template = [
    ["Parameter", "Nilai", "Keterangan"],
    ["NPSN", "20300001", "8 Digit NPSN Sekolah Anda"],
    ["TOKEN", "PASTE_TOKEN_WEBSERVICE_DISINI", "Token dari WebService Dapodik"],
    ["HOST", "127.0.0.1", "IP Komputer yang menjalankan Dapodik Desktop"],
    ["PORT", "5774", "Port WebService Dapodik (default 5774)"]
  ];

  sheet.getRange(1, 1, template.length, 3).setValues(template);
  sheet.getRange(1, 1, 1, 3).setBackground("#0D47A1").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.autoResizeColumns(1, 3);
  SpreadsheetApp.getUi().alert("Sheet 'Config' berhasil disiapkan! Silakan isi NPSN, Token, dan Host komputer Dapodik Anda.");
}

function menuSyncSiswa() {
  var ui = SpreadsheetApp.getUi();
  try {
    var client = Dapodik.fromConfigSheet("Config");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data Siswa") || ss.insertSheet("Data Siswa");
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Sedang menarik data siswa dari Dapodik...", "Proses Sinkronisasi", -1);
    var total = client.syncPesertaDidikToSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().toast("Berhasil menarik " + total + " siswa!", "Selesai", 5);
  } catch (e) {
    ui.alert("Gagal Sinkronisasi Siswa: " + e.message);
  }
}

function menuSyncGtk() {
  var ui = SpreadsheetApp.getUi();
  try {
    var client = Dapodik.fromConfigSheet("Config");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data GTK") || ss.insertSheet("Data GTK");
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Sedang menarik data GTK dari Dapodik...", "Proses Sinkronisasi", -1);
    var total = client.syncGtkToSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().toast("Berhasil menarik " + total + " GTK!", "Selesai", 5);
  } catch (e) {
    ui.alert("Gagal Sinkronisasi GTK: " + e.message);
  }
}

function menuSyncSekolah() {
  var ui = SpreadsheetApp.getUi();
  try {
    var client = Dapodik.fromConfigSheet("Config");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Profil Sekolah") || ss.insertSheet("Profil Sekolah");
    
    var resp = client.getSekolah();
    client.writeToSheet(sheet, resp.rows);
    ui.alert("Profil sekolah berhasil diperbarui!");
  } catch (e) {
    ui.alert("Gagal Sinkronisasi Profil Sekolah: " + e.message);
  }
}

function menuSyncRombel() {
  var ui = SpreadsheetApp.getUi();
  try {
    var client = Dapodik.fromConfigSheet("Config");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data Rombel") || ss.insertSheet("Data Rombel");
    
    var resp = client.getRombonganBelajar();
    client.writeToSheet(sheet, resp.rows);
    ui.alert("Data Rombel berhasil ditarik!");
  } catch (e) {
    ui.alert("Gagal Sinkronisasi Rombel: " + e.message);
  }
}

function menuSyncPrasarana() {
  var ui = SpreadsheetApp.getUi();
  try {
    var client = Dapodik.fromConfigSheet("Config");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data Prasarana") || ss.insertSheet("Data Prasarana");
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Sedang menarik data prasarana dari Dapodik...", "Proses Sinkronisasi", -1);
    var total = client.syncPrasaranaToSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().toast("Berhasil menarik " + total + " data prasarana!", "Selesai", 5);
  } catch (e) {
    ui.alert("Gagal Sinkronisasi Prasarana: " + e.message);
  }
}

