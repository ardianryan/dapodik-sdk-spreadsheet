# 📗 Panduan Integrasi Microsoft Excel VBA (.xlsm)

Toolkit ini memungkinkan operator dan staf sekolah menarik data langsung dari **WebService Dapodik Kemendikdasmen** ke lembar kerja **Microsoft Excel** tanpa perlu mengetik ulang manual.

---

## 🛠️ Cara Memasang di Microsoft Excel (Windows)

### 1. Buka Visual Basic Editor (VBA)
1. Buka Microsoft Excel dan buat file baru, lalu simpan dengan format **Excel Macro-Enabled Workbook (`.xlsm`)**.
2. Tekan kombinasi tombol **`ALT + F11`** pada keyboard untuk membuka VBA Editor.

### 2. Import Modul
1. Di VBA Editor, klik menu **File** ➔ **Import File...** (atau tekan `CTRL + M`).
2. Import file berikut secara berurutan:
   - [`DapodikClient.cls`](DapodikClient.cls) *(Class Module)*
   - [`DapodikSync.bas`](DapodikSync.bas) *(Standard Module)*
   - `JsonConverter.bas` *(JSON Parser library untuk VBA)*

### 3. Menjalankan Macro
1. Kembali ke Excel, tekan **`ALT + F8`** untuk membuka daftar Macro.
2. Jalankan macro **`SetupConfigSheet`** untuk membuat form konfigurasi token.
3. Isi Token, NPSN, dan IP komputer Dapodik pada sheet `Config`.
4. Jalankan macro **`TarikDataSiswa`**!
