/**
 * Digital Lombok - Backend Google Sheets
 * Aksi yang didukung:
 *   GET  ?action=bootstrap  -> master, transactions  (JSONP lewat &callback=)
 *   POST action=saveMaster      data: [{id,n,p,u,k,s}]  (k = jenis peralatan, s = stok; satu data dengan master barang)
 *   POST action=saveTransaction data: {...transaksi lengkap...}
 *   POST action=deleteTransaction data: {id}
 *
 * Sheet yang dipakai: Master, TRANSAKSI_APP, Rekap (+ LOG_ERROR bila ada galat). Sheet lama (TRANSAKSI, DETAIL_ITEM, MASTER_HARGA) tidak disentuh.
 */

// Kosongkan jika script ini dibuat dari menu Extensions > Apps Script di dalam Spreadsheet.
// Isi dengan ID spreadsheet jika script berdiri sendiri.
var SPREADSHEET_ID = "";

var SH_MASTER = "Master";
var SH_TX = "TRANSAKSI_APP"; // sengaja beda dari sheet lama "TRANSAKSI" agar data lama tidak tertimpa
var SH_STOK_LAMA = "Stok"; // sheet lama, hanya dipakai untuk migrasi satu kali
var SH_REKAP = "Rekap";
var SH_SET = "PENGATURAN";

var H_MASTER = ["id", "n", "p", "u", "jenis", "stok"];
var H_TX = ["id", "dibuat", "no", "tgl", "jenis", "paid", "klien", "inst", "acara", "lokasi", "tglacara",
  "disc", "dp", "ppn", "rek", "preset", "syarat", "subtotal", "dpp", "pajak", "total", "dpNominal", "sisa", "items", "tglselesai"];

/** Sheet Transaksi; menambah kolom tglselesai di ujung jika masih susunan lama. */
function txSheet_() {
  var sh = sheet_(SH_TX, H_TX);
  var c = H_TX.length;
  if (String(sh.getRange(1, c).getValue()) !== "tglselesai") sh.getRange(1, c).setValue("tglselesai").setFontWeight("bold");
  return sh;
}

var NUM_COLS_TX = ["dibuat", "paid", "disc", "dp", "subtotal", "dpp", "pajak", "total", "dpNominal", "sisa"];

/* ============================ ENTRY POINT ============================ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === "bootstrap" || !p.action) {
      var stamp = null;
      try { stamp = readStamp_(); } catch (e2) { logErr_("readStamp", e2, ""); }
      out = { ok: true, master: readMaster_(), transactions: readTransactions_(), stamp: stamp };
    } else {
      out = { ok: false, message: "Aksi tidak dikenal: " + p.action };
    }
  } catch (err) {
    logErr_("doGet", err, p.action);
    out = { ok: false, message: String(err && err.message ? err.message : err) };
  }
  return jsonp_(out, p.callback);
}

function doPost(e) {
  var res;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var raw = e && e.parameter && e.parameter.payload;
    if (!raw && e && e.postData) raw = e.postData.contents;
    var payload = JSON.parse(raw || "{}");
    if (payload.action === "saveMaster") res = saveMaster_(payload.data);
    else if (payload.action === "saveStamp") res = saveStamp_(payload.data);
    else if (payload.action === "saveTransaction") res = saveTransaction_(payload.data);
    else if (payload.action === "deleteTransaction") res = deleteTransaction_(payload.data && payload.data.id);
    else res = { ok: false, message: "Aksi tidak dikenal: " + payload.action };
  } catch (err) {
    logErr_("doPost", err, e && e.parameter && e.parameter.payload);
    res = { ok: false, message: String(err && err.message ? err.message : err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return HtmlService.createHtmlOutput("<html><body>" + JSON.stringify(res).replace(/</g, "\\u003c") + "</body></html>");
}

function jsonp_(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb && /^[A-Za-z0-9_$.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ HELPER SHEET ============================ */

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function findSheet_(name) {
  var all = ss_().getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase() === String(name).toLowerCase()) return all[i];
  }
  return null;
}

function logErr_(where, err, raw) {
  try {
    var ss = ss_();
    var sh = findSheet_("LOG_ERROR") || ss.insertSheet("LOG_ERROR");
    sh.appendRow([new Date(), where, String(err && err.stack ? err.stack : err), String(raw || "").slice(0, 300)]);
  } catch (x) {}
}

function sheet_(name, headers) {
  var ss = ss_();
  var sh = findSheet_(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function colIdx_(headers, name) { return headers.indexOf(name) + 1; }

function colLetter_(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function cleanText_(v) {
  if (v === null || v === undefined) return "";
  v = String(v);
  // cegah formula injection
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function num_(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

function readText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return v === null || v === undefined ? "" : String(v);
}

/* ============================ MASTER ============================ */

/** Sheet Master; menambah kolom jenis & stok bila masih susunan lama, dan memindahkan data dari sheet "Stok" lama (satu kali). */
function masterSheet_() {
  var sh = sheet_(SH_MASTER, H_MASTER);
  if (String(sh.getRange(1, 5).getValue()) !== "jenis") {
    sh.getRange(1, 5, 1, 2).setValues([["jenis", "stok"]]).setFontWeight("bold");
    var old = findSheet_(SH_STOK_LAMA);
    var last = sh.getLastRow();
    if (old && old.getLastRow() > 1 && last > 1) {
      var hdr = old.getRange(1, 1, 1, old.getLastColumn()).getValues()[0].map(String);
      var iId = hdr.indexOf("id"), iJ = hdr.indexOf("jenis"), iS = hdr.indexOf("stok");
      var src = old.getRange(2, 1, old.getLastRow() - 1, old.getLastColumn()).getValues();
      var map = {};
      src.forEach(function (r) { if (iId > -1 && r[iId] !== "") map[String(r[iId])] = { k: iJ > -1 ? r[iJ] : "", s: iS > -1 ? r[iS] : "" }; });
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      var vals = ids.map(function (r) { var m = map[String(r[0])]; return m ? [m.k, m.s] : ["", ""]; });
      sh.getRange(2, 5, vals.length, 2).setValues(vals);
    }
    if (old) old.setName("Stok_lama (sudah dipindah ke Master)");
  }
  return sh;
}

function readMaster_() {
  var sh = masterSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, H_MASTER.length).getValues();
  return v.filter(function (r) { return r[0] !== "" || r[1] !== ""; })
    .map(function (r) {
      return { id: String(r[0]), n: String(r[1]), p: num_(r[2]), u: String(r[3] || "unit"),
               k: String(r[4] || ""), s: r[5] === "" ? null : num_(r[5]) };
    });
}

function saveMaster_(data) {
  if (!Array.isArray(data)) throw new Error("Data master tidak valid.");
  var sh = masterSheet_();
  // pertahankan jenis/stok lama bila klien tidak mengirim k/s
  var keep = {};
  readMaster_().forEach(function (m) { keep[m.id] = m; });
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, H_MASTER.length).clearContent();
  if (data.length) {
    var rows = data.map(function (o) {
      var old = keep[String(o.id)] || {};
      var k = (o.k === undefined) ? (old.k || "") : String(o.k || "").trim();
      var s = (o.s === undefined) ? (old.s == null ? "" : old.s)
            : ((o.s === "" || o.s === null) ? "" : Math.max(0, num_(o.s)));
      return [cleanText_(o.id), cleanText_(o.n), num_(o.p), cleanText_(o.u || "unit"), cleanText_(k), s];
    });
    sh.getRange(2, 1, rows.length, H_MASTER.length)
      .setNumberFormats(rows.map(function () { return ["@", "@", "General", "@", "@", "General"]; }));
    sh.getRange(2, 1, rows.length, H_MASTER.length).setValues(rows);
  }
  return { ok: true, message: "Master tersimpan.", count: data.length };
}

/* ============================ TRANSAKSI ============================ */

function readTransactions_() {
  var sh = txSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, H_TX.length).getValues();
  var list = v.filter(function (r) { return r[0] !== ""; }).map(function (r) {
    var o = {};
    H_TX.forEach(function (h, i) {
      var x = r[i];
      if (h === "items") {
        try { o.items = x ? JSON.parse(x) : []; } catch (e) { o.items = []; }
      } else if (h === "ppn") {
        o.ppn = (x === true || String(x).toUpperCase() === "TRUE");
      } else if (NUM_COLS_TX.indexOf(h) > -1) {
        o[h] = num_(x);
      } else {
        o[h] = readText_(x);
      }
    });
    return o;
  });
  list.sort(function (a, b) { return (b.dibuat || 0) - (a.dibuat || 0); });
  return list;
}

function saveTransaction_(d) {
  if (!d || !d.id) throw new Error("Data transaksi tidak valid.");
  var sh = txSheet_();

  var row = H_TX.map(function (h) {
    if (h === "items") return JSON.stringify(Array.isArray(d.items) ? d.items : []);
    if (h === "ppn") return !!d.ppn;
    if (NUM_COLS_TX.indexOf(h) > -1) return num_(d[h]);
    return cleanText_(d[h]);
  });

  // cari baris dengan id yang sama
  var last = sh.getLastRow();
  var target = 0;
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(d.id)) { target = i + 2; break; }
    }
  }
  if (!target) target = last + 1;

  // semua kolom teks disimpan apa adanya (tanggal/nomor/angka-dalam-teks tidak diubah otomatis oleh Sheets)
  var fmts = H_TX.map(function (h) {
    return (NUM_COLS_TX.indexOf(h) > -1 || h === "ppn") ? "General" : "@";
  });
  sh.getRange(target, 1, 1, H_TX.length).setNumberFormats([fmts]);
  sh.getRange(target, 1, 1, H_TX.length).setValues([row]);

  try { rebuildRekap_(); } catch (e) {} // perbarui sheet Rekap
  return { ok: true, message: "Transaksi tersimpan.", id: d.id };
}

function deleteTransaction_(id) {
  id = String(id || "").trim();
  if (!id) throw new Error("ID transaksi tidak valid.");

  var sh = txSheet_();
  var last = sh.getLastRow();
  if (last < 2) throw new Error("Transaksi tidak ditemukan.");

  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var target = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      target = i + 2;
      break;
    }
  }
  if (!target) throw new Error("Transaksi tidak ditemukan.");

  var row = sh.getRange(target, 1, 1, H_TX.length).getValues();
  sh.deleteRow(target);
  try {
    rebuildRekap_();
  } catch (err) {
    sh.insertRowBefore(target);
    sh.getRange(target, 1, 1, H_TX.length).setValues(row);
    throw new Error("Transaksi tidak dihapus karena Rekap gagal diperbarui: " + String(err && err.message ? err.message : err));
  }

  return { ok: true, message: "Transaksi berhasil dihapus.", id: id };
}

/* ============================ STEMPEL (PNG, disimpan di sheet PENGATURAN) ============================ */

var STAMP_CHUNK = 40000; // batas sel Sheets 50.000 karakter, gambar dipecah per baris

function readStamp_() {
  var sh = sheet_(SH_SET, ["key", "value"]);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var v = sh.getRange(2, 1, last - 1, 2).getValues(), parts = [], size = 110, pos = null;
  v.forEach(function (r) {
    var k = String(r[0]);
    if (k === "stempel_ukuran") size = num_(r[1]) || 110;
    else if (k === "stempel_pos") { try { pos = JSON.parse(String(r[1])); } catch (e) {} }
    else if (k.indexOf("stempel_") === 0) parts.push({ i: parseInt(k.slice(8), 10) || 0, s: String(r[1]) });
  });
  parts.sort(function (a, b) { return a.i - b.i; });
  var img = parts.map(function (p) { return p.s; }).join("");
  return img ? { img: img, size: size, pos: pos } : null;
}

function saveStamp_(d) {
  var img = d && d.img ? String(d.img) : "";
  if (img && img.indexOf("data:image/png;base64,") !== 0) throw new Error("Stempel harus berformat PNG.");
  if (img.length > 200000) throw new Error("Gambar stempel terlalu besar.");
  var sh = sheet_(SH_SET, ["key", "value"]);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  if (!img) return { ok: true, message: "Stempel dihapus." };
  var pos = { d: { x: 100, y: -35 }, k: { x: -100, y: -35 } };
  if (d.pos && d.pos.d && d.pos.k) {
    ["d", "k"].forEach(function (t) {
      pos[t] = { x: Math.max(-400, Math.min(400, Math.round(num_(d.pos[t].x)))), y: Math.max(-300, Math.min(300, Math.round(num_(d.pos[t].y)))) };
    });
  }
  var rows = [["stempel_ukuran", String(Math.min(300, Math.max(40, num_(d.size) || 110)))], ["stempel_pos", JSON.stringify(pos)]];
  for (var i = 0, n = 1; i < img.length; i += STAMP_CHUNK, n++) rows.push(["stempel_" + n, img.substr(i, STAMP_CHUNK)]);
  sh.getRange(2, 1, rows.length, 2).setNumberFormats(rows.map(function () { return ["@", "@"]; })); // teks apa adanya
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true, message: "Stempel tersimpan.", chunks: rows.length - 1 };
}

/* ============================ REKAP (nilai, dihitung ulang tiap simpan) ============================ */

/** Menulis ulang sheet Rekap berupa nilai (bukan rumus) dari seluruh transaksi, per bulan tanggal dokumen. */
function rebuildRekap_() {
  var ss = ss_();
  var sh = findSheet_(SH_REKAP) || ss.insertSheet(SH_REKAP);
  var g = {};
  readTransactions_().forEach(function (x) {
    var m = String(x.tgl || "").slice(0, 7);
    if (!m) return;
    var r = g[m] || (g[m] = { nq: 0, vq: 0, ni: 0, vi: 0, k: 0, s: 0 });
    if (x.jenis === "inv") { r.ni++; r.vi += num_(x.total); r.s += num_(x.sisa); }
    else if (x.jenis === "kwt") { r.k += num_(x.paid); }
    else { r.nq++; r.vq += num_(x.total); }
  });
  var rows = Object.keys(g).sort().reverse().map(function (m) {
    var r = g[m];
    return [m, r.nq, r.vq, r.ni, r.vi, r.k, r.s];
  });
  sh.clear();
  sh.getRange(1, 1, 1, 7).setValues([["Bulan", "Jml Penawaran", "Nilai Penawaran", "Jml Invoice", "Nilai Invoice", "Diterima (Kwitansi)", "Sisa Tagihan Invoice"]])
    .setFontWeight("bold");
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat("@");
    sh.getRange(2, 1, rows.length, 7).setValues(rows);
    sh.getRange(2, 3, rows.length, 1).setNumberFormat("#,##0");
    sh.getRange(2, 5, rows.length, 3).setNumberFormat("#,##0");
  }
  sh.setColumnWidths(1, 7, 140);
}

/* ============================ SETUP MANUAL ============================ */

/** Jalankan sekali dari editor Apps Script (pilih fungsi setup > Run) untuk membuat semua sheet dan memberi izin akses. */
function setup() {
  masterSheet_();
  txSheet_();
  rebuildRekap_();
  Logger.log("Sheet Master (termasuk jenis & stok), Transaksi, dan Rekap siap.");
}
