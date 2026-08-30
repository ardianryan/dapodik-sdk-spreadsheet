# 📊 Panduan Integrasi Google Apps Script (Google Sheets)

SDK ini memungkinkan Anda menghubungkan **Google Spreadsheet** secara langsung dengan **WebService Dapodik Kemendikdasmen** (Desktop port 5774).

---

## 🚀 Cara Pemasangan di Google Spreadsheet (5 Menit)

### 1. Buka Apps Script di Google Sheets
1. Buka file Google Spreadsheet baru di browser Anda ([sheets.new](https://sheets.new)).
2. Di menu atas, klik **Extensions (Ekstensi)** ➔ **Apps Script**.

### 2. Salin Kode Script
1. Buat file baru bernama `DapodikClient.gs`, lalu salin seluruh isi dari file [`DapodikClient.gs`](DapodikClient.gs).
2. Buat file baru bernama `CustomMenu.gs`, lalu salin seluruh isi dari file [`CustomMenu.gs`](CustomMenu.gs).
3. Klik tombol **Save 💾 (Simpan)** di Apps Script.

### 3. Menggunakan Menu Dapodik di Google Sheets
1. Kembali ke tab Google Sheets Anda dan lakukan **Refresh / Reload (F5)**.
2. Di baris menu atas akan muncul menu baru: **🎓 Dapodik Kemendikdasmen**.
3. Klik menu **🎓 Dapodik Kemendikdasmen** ➔ **⚙️ Setup Sheet Konfigurasi Token**.
4. Isi NPSN, Token, dan IP Host komputer Dapodik Anda pada sheet `Config`.
5. Sekarang Anda cukup klik **📥 Tarik Semua Siswa** atau **📥 Tarik Semua Guru (GTK)** untuk menarik seluruh data secara otomatis!

---

## ⚡ Contoh Penggunaan Script Mandiri (Custom Macro)

```javascript
function tarikDataKhusus() {
  var client = Dapodik.createClient({
    npsn: "20300001",
    token: "TOKEN_WEBSERVICE_ANDA",
    host: "192.168.1.100", // IP Laptop Dapodik di jaringan lokal
    port: 5774
  });

  // Ambil data profil sekolah
  var sekolah = client.getSekolah();
  Logger.log("Sekolah: " + sekolah.first.nama);

  // Tarik siswa langsung ke sheet aktif
  var sheet = SpreadsheetApp.getActiveSheet();
  client.syncPesertaDidikToSheet(sheet, 100);
}
```
