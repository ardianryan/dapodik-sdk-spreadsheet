# Changelog

Semua perubahan penting pada paket **`dapodik-sdk-spreadsheet`** akan didokumentasikan di file ini.

Format changelog ini mengacu pada [Keep a Changelog](https://keepachangelog.com/id-ID/1.1.0/), dan proyek ini mematuhi [Semantic Versioning](https://semver.org/lang/id/).

---

## [1.0.0] - 2026-08-30

### Ditambahkan
- **Google Apps Script SDK (`google-apps-script/`)**:
  - `DapodikClient.gs`: Core client WebService Dapodik via `UrlFetchApp`.
  - Normalisasi otomatis respons `rows` dan konversi instan ke range Google Sheets dengan pewarnaan header otomatis.
  - Auto-Pagination untuk penarikan data siswa (`syncPesertaDidikToSheet`) dan guru (`syncGtkToSheet`).
  - `CustomMenu.gs`: Menu kustom UI di Google Sheets (`🎓 Dapodik Kemendikdasmen`).
- **Microsoft Excel VBA SDK (`excel-vba/`)**:
  - `DapodikClient.cls`: Class Module HTTP client berbasis `MSXML2.ServerXMLHTTP`.
  - `DapodikSync.bas`: Macro sinkronisasi data siswa ke Excel Worksheet.
- **Kepatuhan Keamanan & UU PDP No. 27/2022**.
