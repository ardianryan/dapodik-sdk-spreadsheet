<p align="center">
  <img src="https://dapo.kemendikdasmen.go.id/assets/logo-dapodik-BZDG7c6h.png" alt="Dapodik Logo" width="140" />
</p>

<h1 align="center">dapodik-sdk-spreadsheet</h1>

<p align="center">
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square" alt="License: MIT" /></a>
  <a href="https://developers.google.com/apps-script"><img src="https://img.shields.io/badge/Google%20Apps%20Script-Ready-4285F4.svg?style=flat-square&logo=google" alt="Google Apps Script Ready" /></a>
  <a href="https://learn.microsoft.com/en-us/office/vba/api/overview/excel"><img src="https://img.shields.io/badge/Excel%20VBA-Ready-107C41.svg?style=flat-square&logo=microsoftexcel" alt="Excel VBA Ready" /></a>
  <a href="https://www.instagram.com/smansagewithai/"><img src="https://img.shields.io/badge/Instagram-@smansagewithai-E4405F.svg?style=flat-square&logo=instagram&logoColor=white" alt="Instagram" /></a>
</p>

<p align="center">
  Toolkit & SDK terpadu untuk integrasi langsung <b>Google Sheets (Google Apps Script)</b> dan <b>Microsoft Excel (VBA Macro)</b> ke <b>WebService Dapodik Kemendikdasmen</b> (port 5774).
</p>

<p align="center">
  Dipublikasikan dan dikelola oleh <b>SMA Negeri 1 Gedeg (<a href="https://www.instagram.com/smansagewithai/">@smansagewithai</a>)</b><br />
  Dikembangkan oleh <b>Ryan Ardian</b>
</p>

---

> [!IMPORTANT]
> ### 📢 Pernyataan Penyangkalan (Disclaimer) & Misi Terbuka
> **`dapodik-sdk-spreadsheet` adalah pustaka *Unofficial* (tidak resmi) dan independen.** Pustaka ini dikembangkan sebagai inisiatif komunitas sumber terbuka (*open-source*) oleh **SMA Negeri 1 Gedeg** dan **Ryan Ardian**, tanpa afiliasi langsung secara struktural dengan Kementerian Pendidikan Dasar dan Menengah (Kemendikdasmen).
>
> **Tujuan & Misi Pengembangan**:
> Pustaka ini diciptakan khusus untuk **memudahkan para operator sekolah, guru, dan staf administrasi** di seluruh Indonesia dalam mengolah data pendidikan di spreadsheet harian (Google Sheets & Microsoft Excel) tanpa perlu proses ekspor-impor manual yang rentan salah ketik (*human error*).
>
> Seluruh hak cipta nama, logo, dan merek dagang **Dapodik (Data Pokok Pendidikan)** adalah milik sah **Kementerian Pendidikan Dasar dan Menengah Republik Indonesia**.

---

## 🏛️ Fitur & Keunggulan

- 📊 **Dukungan Google Spreadsheet (Google Apps Script)**:
  - Menu kustom otomatis di UI Google Sheets: **`🎓 Dapodik Kemendikdasmen`**.
  - Tarik data Siswa, GTK, Rombel, dan Profil Sekolah langsung ke Sheet hanya dengan 1 klik.
  - Penataan tabel dan pewarnaan header otomatis (*auto-styling & freezing header*).
- 📗 **Dukungan Microsoft Excel (VBA Macro `.xlsm`)**:
  - Class Module `DapodikClient.cls` berbasis `MSXML2.ServerXMLHTTP`.
  - Macro otomatis `DapodikSync.bas` untuk sinkronisasi data offline di komputer Windows operator sekolah.
- 🛡️ **Kepatuhan UU PDP No. 27/2022**: Panduan isolasi token dan perlindungan data siswa/guru.

---

## 📂 Struktur Direktori

```text
dapodik-sdk-spreadsheet/
│
├── 📊 google-apps-script/       <-- SDK untuk Google Spreadsheet
│   ├── DapodikClient.gs         <-- Core Client & Sheet Auto-Writer
│   ├── CustomMenu.gs            <-- Menu UI di Google Sheets
│   └── README.md                <-- Panduan instalasi Apps Script
│
└── 📗 excel-vba/                <-- SDK untuk Microsoft Excel
    ├── DapodikClient.cls        <-- Class Module WebService
    ├── DapodikSync.bas          <-- Standard Module Macro
    └── README.md                <-- Panduan import modul ke Excel
```

---

## 🚀 Panduan Cepat Google Spreadsheet

1. Buka Google Spreadsheet baru di browser Anda ([sheets.new](https://sheets.new)).
2. Klik **Extensions** ➔ **Apps Script**.
3. Salin file [`DapodikClient.gs`](google-apps-script/DapodikClient.gs) dan [`CustomMenu.gs`](google-apps-script/CustomMenu.gs).
4. Refresh Google Sheet ➔ Klik menu **🎓 Dapodik Kemendikdasmen** ➔ **⚙️ Setup Sheet Konfigurasi Token**.
5. Masukkan Token dan NPSN, lalu klik **📥 Tarik Semua Siswa**!

---

## 🚀 Panduan Cepat Microsoft Excel VBA

1. Buka Excel, simpan file sebagai **Excel Macro-Enabled Workbook (`.xlsm`)**.
2. Tekan **`ALT + F11`** untuk membuka VBA Editor.
3. Import file [`DapodikClient.cls`](excel-vba/DapodikClient.cls) dan [`DapodikSync.bas`](excel-vba/DapodikSync.bas).
4. Tekan **`ALT + F8`** ➔ Jalankan macro **`SetupConfigSheet`** ➔ Masukkan Token ➔ Jalankan **`TarikDataSiswa`**!

---

## 📄 Lisensi & Kontributor

- **Lisensi**: [MIT License](LICENSE) &copy; 2026 **Ryan Ardian, SMA Negeri 1 Gedeg (smansage)**.
- **Pengembang**: **Ryan Ardian** ([inisaya@ardianryan.com](mailto:inisaya@ardianryan.com)).
- **Inspirasi & Atribusi**: Adaptasi pustaka PHP Dapodik oleh **Ade Reksi Susanto** ([`adereksisusanto/dapodik-api-php`](https://github.com/adereksisusanto/dapodik-api-php)).
