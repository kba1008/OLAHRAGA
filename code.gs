/*************************************************************
 * AtletTrack — Code.gs
 * Pangkalan data TUNGGAL: Google Sheet ini.
 * Cara pasang:
 *  1. Buka Google Sheet baharu > Extensions > Apps Script
 *  2. Tampal fail ini, Save.
 *  3. Run fungsi  setupPangkalanData()  (beri kebenaran).
 *  4. Deploy > New deployment > Web app
 *       Execute as: Me     |    Who has access: Anyone
 *  5. Salin URL /exec dan tampal dalam skrin "Sambung Pangkalan Data" pada app.
 *************************************************************/

var SHEET_GURU = "GURU";
var SHEET_ATLET = "ATLET";
var SHEET_KEHADIRAN = "KEHADIRAN";
var SHEET_ACARA = "ACARA";
var SHEET_JURULATIH = "JURULATIH_ACARA";
var SHEET_PENYERTAAN = "PENYERTAAN";
var PREFIX_REKOD = "REKOD_";

/* ID FOLDER GOOGLE DRIVE untuk simpan gambar atlet.
   Ambil dari URL folder: https://drive.google.com/drive/folders/<ID_INI>
   Biarkan kosong ("") jika mahu skrip cipta folder "GAMBAR ATLET" secara automatik. */
var FOLDER_GAMBAR_ID = "";
var NAMA_FOLDER_GAMBAR = "GAMBAR ATLET";

var ADMIN_EMEL = "admin";
var ADMIN_KATA_LALUAN = "101010";
var MAX_JURULATIH = 10;

var HEADERS = {};
HEADERS[SHEET_GURU] = ["ID", "NAMA PENUH", "EMEL", "KATA LALUAN", "JAWATAN", "SEKOLAH", "NO TELEFON", "PERANAN", "TARIKH DAFTAR"];
HEADERS[SHEET_ATLET] = ["ID", "NAMA PENUH", "NO IC", "JANTINA", "KATEGORI", "SEKOLAH", "GAMBAR (URL)", "CATATAN", "DIDAFTAR OLEH", "TARIKH DAFTAR"];
HEADERS[SHEET_KEHADIRAN] = ["ID", "TARIKH", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "STATUS", "CATATAN", "DICATAT OLEH", "TARIKH & MASA"];
HEADERS[SHEET_ACARA] = ["ACARA", "JENIS", "UNIT", "AKTIF"];
HEADERS[SHEET_JURULATIH] = ["ACARA", "EMEL JURULATIH", "NAMA JURULATIH", "DILANTIK OLEH", "TARIKH LANTIKAN"];
HEADERS[SHEET_PENYERTAAN] = ["ACARA", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "REKOD PERIBADI", "DIMASUKKAN OLEH", "TARIKH"];
var HEADER_REKOD = ["ID", "TARIKH", "MASA", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "KEPUTUSAN", "NILAI", "CATATAN", "DICATAT OLEH", "TARIKH & MASA REKOD"];

var ACARA_LALAI = [
  ["100 M", "MASA", "saat", "YA"],
  ["200 M", "MASA", "saat", "YA"],
  ["400 M", "MASA", "saat", "YA"],
  ["800 M", "MASA", "saat", "YA"],
  ["1500 M", "MASA", "saat", "YA"],
  ["LOMPAT JAUH", "JARAK", "meter", "YA"],
  ["LOMPAT TINGGI", "JARAK", "meter", "YA"],
  ["LONTAR PELURU", "JARAK", "meter", "YA"],
  ["LEMPAR CAKERA", "JARAK", "meter", "YA"],
  ["REJAM LEMBING", "JARAK", "meter", "YA"]
];

/* ---------------- Util ---------------- */
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function kemasSheet(sh, header, warna) {
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  var hr = sh.getRange(1, 1, 1, header.length);
  hr.setFontWeight("bold").setFontColor("#ffffff").setBackground(warna || "#4f2df5")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sh.setRowHeight(1, 34);
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > header.length) sh.deleteColumns(header.length + 1, sh.getMaxColumns() - header.length);
  for (var i = 1; i <= header.length; i++) sh.setColumnWidth(i, 160);
  sh.getRange(1, 1, sh.getMaxRows(), header.length).setFontFamily("Arial");
}

function dapatSheet(nama, header, warna) {
  var s = ss().getSheetByName(nama);
  if (!s) { s = ss().insertSheet(nama); kemasSheet(s, header, warna); }
  else if (s.getLastRow() === 0) kemasSheet(s, header, warna);
  return s;
}

function namaSheetRekod(acara) {
  return (PREFIX_REKOD + String(acara).toUpperCase().replace(/[^A-Z0-9]+/g, "_")).substring(0, 95);
}
function sheetRekod(acara) { return dapatSheet(namaSheetRekod(acara), HEADER_REKOD, "#00a3c4"); }

function baca(nama) {
  var s = ss().getSheetByName(nama);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getDataRange().getValues();
  var h = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, kosong = true;
    for (var j = 0; j < h.length; j++) { o[h[j]] = v[i][j]; if (v[i][j] !== "" && v[i][j] !== null) kosong = false; }
    if (!kosong) out.push(o);
  }
  return out;
}

function idBaharu(prefix, sheetNama) {
  var s = ss().getSheetByName(sheetNama);
  var n = s && s.getLastRow() > 1 ? s.getLastRow() - 1 : 0;
  return prefix + ("000" + (n + 1)).slice(-4) + "-" + String(Date.now()).slice(-4);
}

function nowStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }
function hariIni() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function tarikhStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v || "").substring(0, 10);
}

/* ---------------- Setup ---------------- */
function setupPangkalanData() {
  dapatSheet(SHEET_GURU, HEADERS[SHEET_GURU], "#4f2df5");
  dapatSheet(SHEET_ATLET, HEADERS[SHEET_ATLET], "#00a3c4");
  dapatSheet(SHEET_KEHADIRAN, HEADERS[SHEET_KEHADIRAN], "#0f9d58");
  dapatSheet(SHEET_JURULATIH, HEADERS[SHEET_JURULATIH], "#ff6d00");
  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var sa = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  if (sa.getLastRow() < 2) sa.getRange(2, 1, ACARA_LALAI.length, 4).setValues(ACARA_LALAI);
  baca(SHEET_ACARA).forEach(function (a) { if (a["ACARA"]) sheetRekod(a["ACARA"]); });
  return "Siap. Semua sheet telah dibina.";
}

/* ---------------- Auth ---------------- */
function isAdmin(emel) { return String(emel || "").toLowerCase() === ADMIN_EMEL; }

function cariGuru(emel) {
  var e = String(emel || "").toLowerCase().trim();
  var r = baca(SHEET_GURU).filter(function (g) { return String(g["EMEL"]).toLowerCase().trim() === e; });
  return r.length ? r[0] : null;
}

function daftarGuru(p) {
  dapatSheet(SHEET_GURU, HEADERS[SHEET_GURU], "#4f2df5");
  if (cariGuru(p.emel)) throw new Error("Emel ini telah didaftarkan.");
  var id = idBaharu("G", SHEET_GURU);
  ss().getSheetByName(SHEET_GURU).appendRow([id, p.nama, String(p.emel).toLowerCase().trim(), p.kataLaluan, p.jawatan, p.sekolah, p.telefon, "GURU", nowStr()]);
  return { id: id, nama: p.nama, emel: String(p.emel).toLowerCase().trim(), jawatan: p.jawatan, sekolah: p.sekolah, telefon: p.telefon, peranan: "GURU" };
}

function login(p) {
  var emel = String(p.emel || "").toLowerCase().trim();
  if (emel === ADMIN_EMEL && String(p.kataLaluan) === ADMIN_KATA_LALUAN) {
    return { id: "ADMIN", nama: "Master Admin", emel: ADMIN_EMEL, jawatan: "Master Admin", sekolah: "-", telefon: "-", peranan: "ADMIN" };
  }
  var g = cariGuru(emel);
  if (!g || String(g["KATA LALUAN"]) !== String(p.kataLaluan)) throw new Error("Emel atau kata laluan salah.");
  return { id: g["ID"], nama: g["NAMA PENUH"], emel: g["EMEL"], jawatan: g["JAWATAN"], sekolah: g["SEKOLAH"], telefon: g["NO TELEFON"], peranan: g["PERANAN"] || "GURU" };
}

/* ---------------- Data utama ---------------- */
function semuaData(p) {
  setupPangkalanData();
  var acara = baca(SHEET_ACARA).filter(function (a) { return a["ACARA"]; });
  var rekod = {};
  acara.forEach(function (a) { rekod[a["ACARA"]] = baca(namaSheetRekod(a["ACARA"])); });
  return {
    guru: baca(SHEET_GURU).map(function (g) { return { nama: g["NAMA PENUH"], emel: g["EMEL"], jawatan: g["JAWATAN"], sekolah: g["SEKOLAH"], telefon: g["NO TELEFON"] }; }),
    atlet: baca(SHEET_ATLET),
    kehadiran: baca(SHEET_KEHADIRAN).map(function (k) { k["TARIKH"] = tarikhStr(k["TARIKH"]); return k; }),
    acara: acara,
    jurulatih: baca(SHEET_JURULATIH),
    penyertaan: baca(SHEET_PENYERTAAN),
    rekod: rekod,
    masaPelayan: nowStr()
  };
}

/* ---------------- Gambar Atlet (Google Drive) ---------------- */
function folderGambar() {
  if (FOLDER_GAMBAR_ID) return DriveApp.getFolderById(FOLDER_GAMBAR_ID);
  var it = DriveApp.getFoldersByName(NAMA_FOLDER_GAMBAR);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NAMA_FOLDER_GAMBAR);
}

/* p.gambarBase64 = "data:image/jpeg;base64,...."  ATAU base64 mentah
   p.namaFail     = nama fail pilihan */
function muatNaikGambar(p) {
  if (!p || !p.gambarBase64) throw new Error("Tiada gambar untuk dimuat naik.");
  var data = String(p.gambarBase64);
  var jenis = "image/jpeg";
  var m = data.match(/^data:([^;]+);base64,(.*)$/);
  if (m) { jenis = m[1]; data = m[2]; }
  var nama = (p.namaFail || ("ATLET_" + nowStr().replace(/[^0-9]/g, ""))) + (jenis.indexOf("png") > -1 ? ".png" : ".jpg");
  var blob = Utilities.newBlob(Utilities.base64Decode(data), jenis, nama);
  var fail = folderGambar().createFile(blob);
  try { fail.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { id: fail.getId(), url: "https://drive.google.com/uc?export=view&id=" + fail.getId() };
}

function tambahAtlet(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh menambah atlet baharu.");
  dapatSheet(SHEET_ATLET, HEADERS[SHEET_ATLET], "#00a3c4");
  var id = idBaharu("A", SHEET_ATLET);
  var urlGambar = p.gambar || "";
  if (p.gambarBase64) urlGambar = muatNaikGambar({ gambarBase64: p.gambarBase64, namaFail: id + "_" + String(p.nama).replace(/[^A-Za-z0-9]+/g, "_") }).url;
  ss().getSheetByName(SHEET_ATLET).appendRow([id, p.nama, p.noIc || "", p.jantina || "", p.kategori || "", p.sekolah || "", urlGambar, p.catatan || "", p.olehNama, nowStr()]);
  return { id: id };
}

function kemaskiniAtlet(p) {
  var s = ss().getSheetByName(SHEET_ATLET);
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      var medan = { nama: 1, noIc: 2, jantina: 3, kategori: 4, sekolah: 5, gambar: 6, catatan: 7 };
      Object.keys(medan).forEach(function (k) { if (p[k] !== undefined && p[k] !== null) v[i][medan[k]] = p[k]; });
      v[i][8] = p.olehNama;
      v[i][9] = nowStr();
      s.getRange(i + 1, 1, 1, HEADERS[SHEET_ATLET].length).setValues([v[i].slice(0, HEADERS[SHEET_ATLET].length)]);
      return { ok: true };
    }
  }
  throw new Error("Atlet tidak dijumpai.");
}

function simpanKehadiran(p) {
  dapatSheet(SHEET_KEHADIRAN, HEADERS[SHEET_KEHADIRAN], "#0f9d58");
  var s = ss().getSheetByName(SHEET_KEHADIRAN);
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (tarikhStr(v[i][1]) === p.tarikh && String(v[i][2]) === String(p.atletId)) {
      s.getRange(i + 1, 7, 1, 4).setValues([[p.status, p.catatan || "", p.olehNama, nowStr()]]);
      return { ok: true, dikemaskini: true };
    }
  }
  s.appendRow([idBaharu("K", SHEET_KEHADIRAN), p.tarikh, p.atletId, p.nama, p.kategori || "", p.sekolah || "", p.status, p.catatan || "", p.olehNama, nowStr()]);
  return { ok: true };
}

function lantikJurulatih(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh melantik jurulatih.");
  dapatSheet(SHEET_JURULATIH, HEADERS[SHEET_JURULATIH], "#ff6d00");
  var sedia = baca(SHEET_JURULATIH).filter(function (j) { return j["ACARA"] === p.acara; });
  if (sedia.length >= MAX_JURULATIH) throw new Error("Maksima " + MAX_JURULATIH + " jurulatih bagi setiap acara.");
  var e = String(p.emel).toLowerCase().trim();
  if (sedia.some(function (j) { return String(j["EMEL JURULATIH"]).toLowerCase() === e; })) throw new Error("Jurulatih ini sudah dilantik untuk acara tersebut.");
  ss().getSheetByName(SHEET_JURULATIH).appendRow([p.acara, e, p.nama, p.olehNama, nowStr()]);
  return { ok: true };
}

function buangJurulatih(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh membuang jurulatih.");
  var s = ss().getSheetByName(SHEET_JURULATIH);
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (v[i][0] === p.acara && String(v[i][1]).toLowerCase() === String(p.emel).toLowerCase()) { s.deleteRow(i + 1); return { ok: true }; }
  }
  throw new Error("Rekod jurulatih tidak dijumpai.");
}

function bolehRekod(acara, emel) {
  if (isAdmin(emel)) return true;
  var e = String(emel).toLowerCase();
  return baca(SHEET_JURULATIH).some(function (j) { return j["ACARA"] === acara && String(j["EMEL JURULATIH"]).toLowerCase() === e; });
}

function tambahPenyertaan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh menambah atlet ke acara.");
  var hadir = baca(SHEET_KEHADIRAN).some(function (k) {
    return tarikhStr(k["TARIKH"]) === (p.tarikh || hariIni()) && String(k["ATLET ID"]) === String(p.atletId) && String(k["STATUS"]).toUpperCase() === "HADIR";
  });
  if (!hadir) throw new Error("Kehadiran atlet pada hari ini belum ditanda.");
  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var ada = baca(SHEET_PENYERTAAN).some(function (x) { return x["ACARA"] === p.acara && String(x["ATLET ID"]) === String(p.atletId); });
  if (ada) throw new Error("Atlet sudah berada dalam acara ini.");
  ss().getSheetByName(SHEET_PENYERTAAN).appendRow([p.acara, p.atletId, p.nama, p.kategori || "", p.sekolah || "", p.rekodPeribadi || "", p.olehNama, nowStr()]);
  sheetRekod(p.acara);
  return { ok: true };
}

function simpanRekodLatihan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh merekod.");
  var tarikh = p.tarikh || hariIni();
  var hadir = baca(SHEET_KEHADIRAN).some(function (k) {
    return tarikhStr(k["TARIKH"]) === tarikh && String(k["ATLET ID"]) === String(p.atletId) && String(k["STATUS"]).toUpperCase() === "HADIR";
  });
  if (!hadir) throw new Error("Kehadiran atlet pada " + tarikh + " belum ditanda. Rekod latihan tidak boleh diambil.");
  var s = sheetRekod(p.acara);
  var masa = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  s.appendRow([idBaharu("R", s.getName()), tarikh, masa, p.atletId, p.nama, p.kategori || "", p.sekolah || "", p.keputusan, Number(p.nilai) || "", p.catatan || "", p.olehNama, nowStr()]);
  return { ok: true };
}

function tambahAcara(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh menambah acara.");
  var s = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  var nama = String(p.acara).toUpperCase().trim();
  if (baca(SHEET_ACARA).some(function (a) { return String(a["ACARA"]).toUpperCase() === nama; })) throw new Error("Acara sudah wujud.");
  s.appendRow([nama, p.jenis || "MASA", p.unit || "saat", "YA"]);
  sheetRekod(nama);
  return { ok: true };
}

/* ---------------- Router ---------------- */
var TINDAKAN = {
  ping: function () { return { ok: true, masa: nowStr() }; },
  setup: function () { return { mesej: setupPangkalanData() }; },
  daftar: daftarGuru,
  login: login,
  data: semuaData,
  muatNaikGambar: muatNaikGambar,
  tambahAtlet: tambahAtlet,
  kemaskiniAtlet: kemaskiniAtlet,
  kehadiran: simpanKehadiran,
  lantikJurulatih: lantikJurulatih,
  buangJurulatih: buangJurulatih,
  tambahPenyertaan: tambahPenyertaan,
  rekod: simpanRekodLatihan,
  tambahAcara: tambahAcara
};

function balas(obj, callback) {
  var teks = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + teks + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(teks).setMimeType(ContentService.MimeType.JSON);
}

/* Klasifikasi ralat supaya teknision mudah mengesan punca masalah */
function kodRalat(mesej) {
  var m = String(mesej || "").toLowerCase();
  if (m.indexOf("tindakan tidak sah") >= 0) return "DB-100";           // aksi tidak dikenali
  if (m.indexOf("kata laluan") >= 0 || m.indexOf("emel") >= 0) return "DB-200"; // pengesahan
  if (m.indexOf("hanya") >= 0 || m.indexOf("maksima") >= 0) return "DB-300";    // kebenaran / had
  if (m.indexOf("kehadiran") >= 0) return "DB-400";                    // peraturan kehadiran
  if (m.indexOf("tidak dijumpai") >= 0 || m.indexOf("sudah") >= 0) return "DB-500"; // data
  if (m.indexOf("sheet") >= 0 || m.indexOf("range") >= 0 || m.indexOf("lajur") >= 0) return "DB-006";
  if (m.indexOf("permission") >= 0 || m.indexOf("authoriz") >= 0) return "DB-007";
  if (m.indexOf("lock") >= 0 || m.indexOf("timeout") >= 0) return "DB-002";
  return "DB-004";
}

function logRalat(payload, err) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName("LOG RALAT");
    if (!sh) {
      sh = ss.insertSheet("LOG RALAT");
      sh.appendRow(["TARIKH & MASA", "KOD", "AKSI", "MESEJ", "OLEH", "BUTIRAN"]);
      sh.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#1f2937").setFontColor("#ffffff");
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), kodRalat(err && err.message), (payload && payload.action) || "-",
      String((err && err.message) || err), (payload && payload.olehEmel) || "-",
      String((err && err.stack) || "").slice(0, 900)]);
  } catch (x) { /* jangan biarkan log gagal menghalang balasan */ }
}

function proses(payload, callback) {
  var mula = new Date().getTime();
  try {
    var fn = TINDAKAN[payload.action];
    if (!fn) throw new Error("Tindakan tidak sah: " + payload.action);
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try { return balas({ ok: true, data: fn(payload), ms: new Date().getTime() - mula }, callback); }
    finally { lock.releaseLock(); }
  } catch (err) {
    logRalat(payload, err);
    return balas({
      ok: false,
      kod: kodRalat(err && err.message),
      error: String((err && err.message) || err),
      aksi: (payload && payload.action) || "-",
      masa: new Date().toISOString(),
      butiran: String((err && err.stack) || "").slice(0, 600)
    }, callback);
  }
}

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  if (p.payload) { try { p = JSON.parse(p.payload); } catch (x) {} }
  if (!p.action) p.action = "ping";
  return proses(p, (e && e.parameter && e.parameter.callback) || null);
}

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) { p = (e && e.parameter) || {}; }
  return proses(p, null);
}
