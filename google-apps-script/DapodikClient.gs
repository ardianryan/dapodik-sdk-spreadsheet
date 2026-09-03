/**
 * @file DapodikClient.gs
 * @description SDK Google Apps Script untuk WebService Dapodik Kemendikdasmen.
 * @author Ryan Ardian <inisaya@ardianryan.com>
 * @publisher SMA Negeri 1 Gedeg (@smansagewithai)
 * @license MIT
 */

/**
 * Factory class untuk inisialisasi DapodikClient di Google Apps Script
 */
var Dapodik = {
  /**
   * Membuat instance DapodikClient
   * @param {Object} config - Konfigurasi { npsn, token, host, port, baseUrl, timeout }
   * @returns {DapodikClient}
   */
  createClient: function(config) {
    return new DapodikClient(config);
  },

  /**
   * Membuat client dari Sheet konfigurasi aktif bernama 'Config'
   * @returns {DapodikClient}
   */
  fromConfigSheet: function(sheetName) {
    sheetName = sheetName || "Config";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Sheet '" + sheetName + "' tidak ditemukan. Buat sheet dengan kolom Pengaturan dan Nilai.");
    }
    var data = sheet.getDataRange().getValues();
    var cfg = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      var val = String(data[i][1]).trim();
      if (key === "npsn") cfg.npsn = val;
      if (key === "token") cfg.token = val;
      if (key === "host" || key === "ip") cfg.host = val;
      if (key === "port") cfg.port = val;
      if (key === "base_url" || key === "baseurl") cfg.baseUrl = val;
    }
    return new DapodikClient(cfg);
  }
};

/**
 * Class DapodikClient untuk Google Apps Script
 */
function DapodikClient(config) {
  if (!config || !config.npsn || !config.token) {
    throw new Error("DapodikClient Error: 'npsn' dan 'token' wajib diisi.");
  }

  var npsnStr = String(config.npsn).trim();
  var tokenStr = String(config.token).trim();

  if (/[\r\n]/.test(npsnStr)) {
    throw new Error("DapodikClient Error: NPSN tidak boleh mengandung karakter newline.");
  }
  if (/[\r\n]/.test(tokenStr)) {
    throw new Error("DapodikClient Error: Token tidak boleh mengandung karakter newline (CRLF injection prevention).");
  }

  this.npsn = npsnStr;
  this.token = tokenStr;
  this.timeout = config.timeout || 30000;

  if (config.baseUrl) {
    this.baseUrl = config.baseUrl.replace(/\/+$/, "");
  } else {
    var host = (config.host || "127.0.0.1").trim();
    if (!host.match(/^https?:\/\//i)) {
      host = "http://" + host;
    }
    var port = config.port || 5774;
    this.baseUrl = host + ":" + port + "/WebService";
  }
}

/**
 * Mengirim HTTP Request ke WebService Dapodik via UrlFetchApp
 */
DapodikClient.prototype.request = function(method, endpoint, params, body) {
  var cleanEndpoint = endpoint.replace(/^\/+/, "");
  if (cleanEndpoint.indexOf("..") !== -1 || cleanEndpoint.indexOf("\\") !== -1) {
    throw new Error("DapodikClient Error: Endpoint tidak valid (path traversal detected).");
  }
  var queryObj = { npsn: this.npsn };
  
  if (params && typeof params === "object") {
    for (var key in params) {
      if (params.hasOwnProperty(key) && params[key] !== null && params[key] !== undefined) {
        queryObj[key] = params[key];
      }
    }
  }

  var queryString = Object.keys(queryObj)
    .map(function(k) { return encodeURIComponent(k) + "=" + encodeURIComponent(queryObj[k]); })
    .join("&");

  var url = this.baseUrl + "/" + cleanEndpoint + "?" + queryString;

  var options = {
    method: method.toLowerCase(),
    headers: {
      "Authorization": "Bearer " + this.token,
      "Accept": "application/json, text/plain, */*",
      "User-Agent": "dapodik-sdk-apps-script/1.0.0"
    },
    muteHttpExceptions: true,
    validateHttpsCertificates: false
  };

  if (body) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(body);
  }

  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    throw new Error("Gagal terhubung ke WebService Dapodik (" + this.baseUrl + "): " + e.message);
  }

  var statusCode = response.getResponseCode();
  var rawText = response.getContentText();

  if (statusCode === 401 || statusCode === 403) {
    throw new Error("Autentikasi Gagal (" + statusCode + "): Periksa token atau pastikan IP client sudah di-whitelist di WebService Dapodik.");
  }

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Dapodik HTTP Error " + statusCode + ": " + rawText.substring(0, 200));
  }

  var parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (e) {
    throw new Error("Respons bukan JSON yang valid dari Dapodik: " + rawText.substring(0, 200));
  }

  // Normalisasi rows
  var rows = [];
  if (Array.isArray(parsed)) {
    rows = parsed;
  } else if (parsed && typeof parsed === "object") {
    if (parsed.rows) {
      if (Array.isArray(parsed.rows)) {
        rows = parsed.rows;
      } else if (typeof parsed.rows === "object") {
        rows = [parsed.rows]; // Single object normalization (getSekolah)
      }
    } else if (Array.isArray(parsed.data)) {
      rows = parsed.data;
    } else {
      rows = [parsed];
    }
  }

  return {
    rows: rows,
    raw: parsed,
    count: rows.length,
    first: rows.length > 0 ? rows[0] : null
  };
};

// =============================================================================
// Endpoint API
// =============================================================================

DapodikClient.prototype.getSekolah = function(params) {
  var p = typeof params === "string" ? { semester_id: params } : (params || {});
  return this.request("get", "getSekolah", p);
};

DapodikClient.prototype.getPengguna = function(params) {
  return this.request("get", "getPengguna", params);
};

DapodikClient.prototype.getGtk = function(page, limit, params) {
  var p = params || {};
  if (page) p.page = page;
  if (limit) p.limit = limit;
  return this.request("get", "getGtk", p);
};

DapodikClient.prototype.getRombonganBelajar = function(semesterId, params) {
  var p = params || {};
  if (semesterId) p.semester_id = semesterId;
  return this.request("get", "getRombonganBelajar", p);
};

DapodikClient.prototype.getPesertaDidik = function(page, limit, params) {
  var p = params || {};
  if (page) p.page = page;
  if (limit) p.limit = limit;
  return this.request("get", "getPesertaDidik", p);
};

DapodikClient.prototype.getMataPelajaran = function(semesterId, params) {
  var p = params || {};
  if (semesterId) p.semester_id = semesterId;
  return this.request("get", "getMataPelajaran", p);
};

DapodikClient.prototype.getMatevNilai = function(semesterId, params) {
  var p = params || {};
  if (semesterId) p.semester_id = semesterId;
  return this.request("get", "getMatevNilai", p);
};

// Aliases
DapodikClient.prototype.sekolah = DapodikClient.prototype.getSekolah;
DapodikClient.prototype.pengguna = DapodikClient.prototype.getPengguna;
DapodikClient.prototype.gtk = DapodikClient.prototype.getGtk;
DapodikClient.prototype.rombel = DapodikClient.prototype.getRombonganBelajar;
DapodikClient.prototype.pd = DapodikClient.prototype.getPesertaDidik;
DapodikClient.prototype.mataPelajaran = DapodikClient.prototype.getMataPelajaran;
DapodikClient.prototype.matevNilai = DapodikClient.prototype.getMatevNilai;

// =============================================================================
// Sinkronisasi Langsung ke Google Sheet
// =============================================================================

/**
 * Menulis array of objects ke Sheet dengan header otomatis
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target Sheet
 * @param {Array<Object>} rows - Data baris
 * @param {Array<string>} [selectedColumns] - Kolom yang dipilih (opsional)
 */
DapodikClient.prototype.writeToSheet = function(sheet, rows, selectedColumns) {
  if (!sheet) throw new Error("Sheet target tidak valid.");
  if (!rows || rows.length === 0) {
    sheet.getRange("A1").setValue("Tidak ada data ditemukan dari Dapodik.");
    return;
  }

  // Tentukan kolom header
  var headers = selectedColumns;
  if (!headers || headers.length === 0) {
    headers = [];
    for (var i = 0; i < rows.length; i++) {
      for (var key in rows[i]) {
        if (rows[i].hasOwnProperty(key) && headers.indexOf(key) === -1) {
          headers.push(key);
        }
      }
    }
  }

  // Siapkan matriks 2D
  var output = [];
  output.push(headers); // Baris 1: Header

  for (var r = 0; r < rows.length; r++) {
    var rowData = [];
    for (var c = 0; c < headers.length; c++) {
      var val = rows[r][headers[c]];
      if (val === null || val === undefined) {
        rowData.push("");
      } else if (typeof val === "object") {
        rowData.push(JSON.stringify(val));
      } else {
        rowData.push(String(val));
      }
    }
    output.push(rowData);
  }

  // Bersihkan sheet lalu tulis seluruh matriks
  sheet.clear();
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  // Format header
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1B5E20").setFontColor("#FFFFFF").setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
};

/**
 * Tarik seluruh siswa dan tulis langsung ke Sheet
 */
DapodikClient.prototype.syncPesertaDidikToSheet = function(sheet, limit) {
  limit = limit || 100;
  var allRows = [];
  var page = 1;

  while (true) {
    var resp = this.getPesertaDidik(page, limit);
    if (!resp.rows || resp.rows.length === 0) break;
    allRows = allRows.concat(resp.rows);
    if (resp.rows.length < limit) break;
    page++;
    Utilities.sleep(100);
  }

  this.writeToSheet(sheet, allRows);
  return allRows.length;
};

/**
 * Tarik seluruh GTK dan tulis langsung ke Sheet
 */
DapodikClient.prototype.syncGtkToSheet = function(sheet, limit) {
  limit = limit || 100;
  var allRows = [];
  var page = 1;

  while (true) {
    var resp = this.getGtk(page, limit);
    if (!resp.rows || resp.rows.length === 0) break;
    allRows = allRows.concat(resp.rows);
    if (resp.rows.length < limit) break;
    page++;
    Utilities.sleep(100);
  }

  this.writeToSheet(sheet, allRows);
  return allRows.length;
};
