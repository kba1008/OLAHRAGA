/*************************************************************
 * AtletTraning — Code.gs  (versi kemaskini)
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
var SHEET_FAIL = "FAIL_ATLET";
var SHEET_KEJOHANAN = "REKOD_KEJOHANAN";
var SHEET_KAT_GUGUR = "KATEGORI_GUGUR";
var SHEET_BMI = "BMI";
var PREFIX_REKOD = "REKOD_";

/* ID Google Sheet UTAMA (pangkalan data). Skrip akan buka sheet ini terus,
   walaupun skrip dijalankan sebagai skrip standalone (tidak terikat pada sheet). */
var SHEET_ID = "1Y2sYo-8PqhCkVP9W0RbYfXnQWx4efKsP_6-veIm7dxM";

/* ID FOLDER GOOGLE DRIVE untuk simpan gambar atlet.
   Ambil dari URL folder: https://drive.google.com/drive/folders/<ID_INI>
   Biarkan kosong ("") jika mahu skrip cipta folder "GAMBAR ATLET" secara automatik. */
var FOLDER_GAMBAR_ID = "1Nz0S__dRbA4vP4Ca0xBRhpdPNUj4KVOf";
var NAMA_FOLDER_GAMBAR = "GAMBAR ATLET";
var NAMA_FOLDER_FAIL = "FAIL ATLET";
var MAX_FAIL_ATLET = 5;

var ADMIN_EMEL = "admin";
var ADMIN_KATA_LALUAN = "101010";
var MAX_JURULATIH = 10;

var HEADERS = {};
HEADERS[SHEET_GURU] = ["ID", "NAMA PENUH", "EMEL", "KATA LALUAN", "JAWATAN", "SEKOLAH", "NO TELEFON", "PERANAN", "TARIKH DAFTAR"];
HEADERS[SHEET_ATLET] = ["ID", "NAMA PENUH", "NO IC", "JANTINA", "KATEGORI", "SEKOLAH", "GAMBAR (URL)", "CATATAN", "DIDAFTAR OLEH", "TARIKH DAFTAR"];
HEADERS[SHEET_KEHADIRAN] = ["ID", "TARIKH", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "STATUS", "CATATAN", "DICATAT OLEH", "TARIKH & MASA"];
HEADERS[SHEET_ACARA] = ["ACARA", "JENIS", "UNIT", "MOD", "SUSUNAN", "AKTIF"];
HEADERS[SHEET_JURULATIH] = ["ACARA", "EMEL JURULATIH", "NAMA JURULATIH", "DILANTIK OLEH", "TARIKH LANTIKAN"];
HEADERS[SHEET_PENYERTAAN] = ["ACARA", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "REKOD PERIBADI", "DIMASUKKAN OLEH", "TARIKH", "SUSUNAN"];
HEADERS[SHEET_KEJOHANAN] = ["ID", "ACARA", "KATEGORI", "NAMA KEJOHANAN", "TAHUN", "PEMEGANG REKOD", "NILAI", "KEPUTUSAN", "CATATAN", "DITETAPKAN OLEH", "TARIKH & MASA"];
HEADERS[SHEET_KAT_GUGUR] = ["ACARA", "KATEGORI", "STATUS", "OLEH", "TARIKH & MASA"];
HEADERS[SHEET_BMI] = ["ID", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "TARIKH", "TINGGI (CM)", "BERAT (KG)", "BMI", "STATUS", "CATATAN", "DICATAT OLEH", "TARIKH & MASA"];
HEADERS[SHEET_FAIL] = ["ID", "ATLET ID", "NAMA FAIL", "JENIS", "SAIZ (BYTES)", "URL", "DRIVE ID", "DIMUAT NAIK OLEH", "TARIKH & MASA"];
var HEADER_REKOD = ["ID", "TARIKH", "MASA", "ATLET ID", "NAMA ATLET", "KATEGORI", "SEKOLAH", "KEPUTUSAN", "NILAI", "CATATAN", "DICATAT OLEH", "TARIKH & MASA REKOD"];

var ACARA_LALAI = [
  ["100 M", "MASA", "saat", "INDIVIDU", 1, "YA"],
  ["200 M", "MASA", "saat", "INDIVIDU", 2, "YA"],
  ["400 M", "MASA", "saat", "INDIVIDU", 3, "YA"],
  ["800 M", "MASA", "saat", "INDIVIDU", 4, "YA"],
  ["1500 M", "MASA", "saat", "INDIVIDU", 5, "YA"],
  ["LOMPAT JAUH", "JARAK", "meter", "INDIVIDU", 6, "YA"],
  ["LOMPAT TINGGI", "JARAK", "meter", "INDIVIDU", 7, "YA"],
  ["LONTAR PELURU", "JARAK", "meter", "INDIVIDU", 8, "YA"],
  ["LEMPAR CAKERA", "JARAK", "meter", "INDIVIDU", 9, "YA"],
  ["REJAM LEMBING", "JARAK", "meter", "INDIVIDU", 10, "YA"]
];

/* ---------------- Util ---------------- */
var __SS_CACHE = null;
function ss() {
  if (__SS_CACHE) return __SS_CACHE;
  var a = null;
  try { a = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!a && SHEET_ID) a = SpreadsheetApp.openById(SHEET_ID);
  if (!a) throw new Error("Tidak dapat buka Google Sheet. Semak SHEET_ID atau buka Apps Script dari sheet yang betul.");
  __SS_CACHE = a;
  return a;
}

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

/* Tukar sel Date kepada teks yang betul supaya tidak keluar 1899-12-30T...Z */
var TZ_MY = "Asia/Kuala_Lumpur";
function tz_() { return TZ_MY; }

function nilaiSelamat(header, v) {
  if (!(v instanceof Date)) return v;
  var tz = tz_();
  var hh = String(header || "").toUpperCase();
  if (hh === "MASA") return Utilities.formatDate(v, tz, "HH:mm:ss");
  if (hh === "TARIKH") return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  return Utilities.formatDate(v, tz, "yyyy-MM-dd HH:mm:ss");
}

function baca(nama) {
  var s = ss().getSheetByName(nama);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getDataRange().getValues();
  var h = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, kosong = true;
    for (var j = 0; j < h.length; j++) { o[h[j]] = nilaiSelamat(h[j], v[i][j]); if (v[i][j] !== "" && v[i][j] !== null) kosong = false; }
    if (!kosong) out.push(o);
  }
  return out;
}

function idBaharu(prefix, sheetNama) {
  var s = ss().getSheetByName(sheetNama);
  var n = s && s.getLastRow() > 1 ? s.getLastRow() - 1 : 0;
  return prefix + ("000" + (n + 1)).slice(-4) + "-" + String(Date.now()).slice(-4);
}

function nowStr() { return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm:ss"); }
function hariIni() { return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd"); }
function tarikhStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), "yyyy-MM-dd");
  return String(v || "").substring(0, 10);
}

/* ---------------- Setup ---------------- */
var _SETUP_DONE = false;
function setupPangkalanData() {
  dapatSheet(SHEET_GURU, HEADERS[SHEET_GURU], "#4f2df5");
  dapatSheet(SHEET_ATLET, HEADERS[SHEET_ATLET], "#00a3c4");
  dapatSheet(SHEET_KEHADIRAN, HEADERS[SHEET_KEHADIRAN], "#0f9d58");
  dapatSheet(SHEET_JURULATIH, HEADERS[SHEET_JURULATIH], "#ff6d00");
  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  dapatSheet(SHEET_FAIL, HEADERS[SHEET_FAIL], "#0aa5d6");
  dapatSheet(SHEET_KEJOHANAN, HEADERS[SHEET_KEJOHANAN], "#b45309");
  dapatSheet(SHEET_KAT_GUGUR, HEADERS[SHEET_KAT_GUGUR], "#b42318");
  dapatSheet(SHEET_BMI, HEADERS[SHEET_BMI], "#0f766e");
  try { pastikanKolumTinggi(); } catch (e) {}
  var sa = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  if (sa.getLastRow() < 2) sa.getRange(2, 1, ACARA_LALAI.length, 6).setValues(ACARA_LALAI);
  baca(SHEET_ACARA).forEach(function (a) { if (a["ACARA"]) sheetRekod(a["ACARA"]); });
  _SETUP_DONE = true;
  return "Siap. Semua sheet telah dibina.";
}
function setupSekaliSahaja() {
  if (_SETUP_DONE) return;
  // sanity check only
  if (!ss().getSheetByName(SHEET_ATLET) || !ss().getSheetByName(SHEET_FAIL)) setupPangkalanData();
  else _SETUP_DONE = true;
}

/* ---------------- Auth ---------------- */
function isMasterAdmin(emel) { return String(emel || "").toLowerCase().trim() === ADMIN_EMEL; }
/* Sub Admin — guru yang dilantik oleh Master Admin (PERANAN = SUBADMIN) */
function isSubAdmin(emel) {
  var e = String(emel || "").toLowerCase().trim();
  if (!e || e === ADMIN_EMEL) return false;
  var g = cariGuru(e);
  return !!g && String(g["PERANAN"] || "").toUpperCase().trim() === "SUBADMIN";
}
/* Sub Admin mempunyai kuasa penuh sama seperti Master Admin (kecuali padam Master Admin) */
function isAdmin(emel) { return isMasterAdmin(emel) || isSubAdmin(emel); }

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
  setupSekaliSahaja();
  try { pastikanKolumAcara(); } catch (e) {}
  try { pastikanKolumPenyertaan(); } catch (e) {}
  try { pastikanKolumTinggi(); } catch (e) {}
  var acara = baca(SHEET_ACARA).filter(function (a) { return a["ACARA"]; });
  acara.sort(function (a, b) {
    var x = Number(a["SUSUNAN"] || 0) || 9999, y = Number(b["SUSUNAN"] || 0) || 9999;
    return x - y;
  });
  var rekod = {};
  acara.forEach(function (a) { rekod[a["ACARA"]] = baca(namaSheetRekod(a["ACARA"])); });
  return {
    guru: baca(SHEET_GURU).map(function (g) { return { nama: g["NAMA PENUH"], emel: g["EMEL"], jawatan: g["JAWATAN"], sekolah: g["SEKOLAH"], telefon: g["NO TELEFON"], peranan: String(g["PERANAN"] || "GURU").toUpperCase() }; }),
    atlet: baca(SHEET_ATLET).map(function (a) { a["GAMBAR (URL)"] = normalGambar(a["GAMBAR (URL)"]); return a; }),
    kehadiran: baca(SHEET_KEHADIRAN).map(function (k) { k["TARIKH"] = tarikhStr(k["TARIKH"]); return k; }),
    acara: acara,
    jurulatih: baca(SHEET_JURULATIH),
    penyertaan: baca(SHEET_PENYERTAAN),
    rekod: rekod,
    failAtlet: baca(SHEET_FAIL),
    kejohanan: baca(SHEET_KEJOHANAN),
    katGugur: baca(SHEET_KAT_GUGUR),
    bmi: baca(SHEET_BMI).map(function (b) { b["TARIKH"] = tarikhStr(b["TARIKH"]); return b; }),
    masaPelayan: nowStr()
  };
}

/* ---------------- Folder Google Drive ---------------- */
var FOLDER_FAIL_ID = ""; /* boleh diisi dengan ID folder "FAIL ATLET" jika mahu tetap */

function cariAtauCiptaFolder(nama, indukId) {
  /* Cuba cari dalam folder induk dahulu (kurang kebenaran diperlukan),
     kemudian barulah carian global, akhir sekali cipta baharu. */
  if (indukId) {
    try {
      var induk = DriveApp.getFolderById(indukId);
      var it0 = induk.getFoldersByName(nama);
      if (it0.hasNext()) return it0.next();
      return induk.createFolder(nama);
    } catch (e) {}
  }
  try {
    var it = DriveApp.getFoldersByName(nama);
    if (it.hasNext()) return it.next();
  } catch (e) {
    throw new Error("Kebenaran Google Drive belum diberikan. Buka Apps Script > pilih fungsi authorizeAll > Run > benarkan akses. Butiran: " + e.message);
  }
  return DriveApp.createFolder(nama);
}

function folderGambar() {
  if (FOLDER_GAMBAR_ID) { try { return DriveApp.getFolderById(FOLDER_GAMBAR_ID); } catch (e) {} }
  return cariAtauCiptaFolder(NAMA_FOLDER_GAMBAR, null);
}

function folderFail() {
  if (FOLDER_FAIL_ID) { try { return DriveApp.getFolderById(FOLDER_FAIL_ID); } catch (e) {} }
  /* Letak folder FAIL ATLET bersebelahan folder GAMBAR ATLET jika ada */
  var indukId = null;
  if (FOLDER_GAMBAR_ID) {
    try {
      var par = DriveApp.getFolderById(FOLDER_GAMBAR_ID).getParents();
      if (par.hasNext()) indukId = par.next().getId();
    } catch (e) {}
  }
  return cariAtauCiptaFolder(NAMA_FOLDER_FAIL, indukId);
}

/* ====== JALANKAN FUNGSI INI SEKALI UNTUK BERI SEMUA KEBENARAN ======
   Apps Script > pilih "authorizeAll" > Run > Benarkan (Allow) semua.
   Kemudian: Deploy > Manage deployments > Edit > New version > Deploy. */
function authorizeAll() {
  var log = [];
  try { log.push("Emel: " + Session.getEffectiveUser().getEmail()); } catch (e) { log.push("Emel: gagal - " + e.message); }
  try { log.push("Sheet: " + ss().getName()); } catch (e) { log.push("Sheet: GAGAL - " + e.message); }
  try { setupPangkalanData(); log.push("Setup sheet: OK"); } catch (e) { log.push("Setup sheet: GAGAL - " + e.message); }
  try { var fg = folderGambar(); log.push("Folder gambar: " + fg.getName() + " (" + fg.getId() + ")"); }
  catch (e) { log.push("Folder gambar: GAGAL - " + e.message); }
  try { var ff = folderFail(); log.push("Folder fail: " + ff.getName() + " (" + ff.getId() + ")"); }
  catch (e) { log.push("Folder fail: GAGAL - " + e.message); }
  try {
    var uji = folderFail().createFile(Utilities.newBlob("ujian kebenaran", "text/plain", "UJIAN_AUTHORIZE.txt"));
    uji.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    uji.setTrashed(true);
    log.push("Ujian cipta & padam fail: OK");
  } catch (e) { log.push("Ujian cipta fail: GAGAL - " + e.message); }
  try { UrlFetchApp.fetch("https://www.google.com/generate_204"); log.push("Capaian luar: OK"); }
  catch (e) { log.push("Capaian luar: GAGAL - " + e.message); }
  var hasil = log.join("\n");
  Logger.log(hasil);
  return hasil;
}

/* p.gambarBase64 = "data:image/jpeg;base64,...."  ATAU base64 mentah
   p.namaFail     = nama fail pilihan
   NOTA: Google telah menyekat "uc?export=view" untuk <img>. Kita guna
   lh3.googleusercontent.com/d/<ID> yang boleh dipapar terus dalam app. */
function urlGambarDrive(id) { return "https://lh3.googleusercontent.com/d/" + id + "=w600"; }

/* Tukar URL Drive lama (uc?export=view / open?id / /file/d/ID/view) kepada
   bentuk lh3 yang boleh dipaparkan dalam <img>. */
function normalGambar(u) {
  var s = String(u || "").trim();
  if (!s) return "";
  if (s.indexOf("lh3.googleusercontent.com") > -1) return s;
  var m = s.match(/[-\w]{25,}/);
  if (s.indexOf("drive.google.com") > -1 && m) return urlGambarDrive(m[0]);
  return s;
}

/* Jalankan sekali dalam Apps Script untuk membaiki URL gambar atlet yang lama. */
function baikiUrlGambar() {
  var s = ss().getSheetByName(SHEET_ATLET);
  if (!s || s.getLastRow() < 2) return "Tiada atlet.";
  var n = s.getLastRow() - 1;
  var rng = s.getRange(2, 7, n, 1), v = rng.getValues(), ubah = 0;
  for (var i = 0; i < v.length; i++) {
    var baru = normalGambar(v[i][0]);
    if (baru !== v[i][0]) { v[i][0] = baru; ubah++; }
  }
  rng.setValues(v);
  return "Siap. " + ubah + " URL gambar dibaiki.";
}

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
  return { id: fail.getId(), url: urlGambarDrive(fail.getId()) };
}


/* ---------------- Fail Lampiran Atlet ---------------- */
function senaraiFailAtlet(atletId) {
  return baca(SHEET_FAIL).filter(function (f) { return String(f["ATLET ID"]) === String(atletId); });
}

function muatNaikFailAtlet(p) {
  if (!p || !p.atletId) throw new Error("ID atlet diperlukan.");
  if (!p.dataBase64) throw new Error("Tiada fail untuk dimuat naik.");
  var sedia = senaraiFailAtlet(p.atletId);
  if (sedia.length >= MAX_FAIL_ATLET) throw new Error("Maksima " + MAX_FAIL_ATLET + " fail bagi setiap atlet. Sila padam fail lama dahulu.");
  var data = String(p.dataBase64), jenis = p.jenis || "application/octet-stream", m = data.match(/^data:([^;]+);base64,(.*)$/);
  if (m) { jenis = m[1]; data = m[2]; }
  var bytes = Utilities.base64Decode(data);
  var nama = String(p.namaFail || "fail").replace(/[\\/:*?"<>|]/g, "_").slice(0, 120);
  var blob = Utilities.newBlob(bytes, jenis, p.atletId + "__" + nama);
  var fail = folderFail().createFile(blob);
  try { fail.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var url = "https://drive.google.com/uc?export=view&id=" + fail.getId();
  var id = idBaharu("F", SHEET_FAIL);
  dapatSheet(SHEET_FAIL, HEADERS[SHEET_FAIL], "#0aa5d6");
  ss().getSheetByName(SHEET_FAIL).appendRow([id, p.atletId, nama, jenis, bytes.length, url, fail.getId(), p.olehNama || "", nowStr()]);
  return { id: id, "ATLET ID": p.atletId, "NAMA FAIL": nama, "JENIS": jenis, "SAIZ (BYTES)": bytes.length, "URL": url, "DRIVE ID": fail.getId() };
}

function padamFailAtlet(p) {
  var s = ss().getSheetByName(SHEET_FAIL);
  if (!s) throw new Error("Tiada rekod fail.");
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      try { DriveApp.getFileById(String(v[i][6])).setTrashed(true); } catch (e) {}
      s.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error("Fail tidak dijumpai.");
}

/* Jurulatih yang DILANTIK (ada dalam JURULATIH_ACARA) atau Master Admin */
function isJurulatihDilantik(emel) {
  var e = String(emel || "").toLowerCase().trim();
  if (!e) return false;
  return baca(SHEET_JURULATIH).some(function (j) { return String(j["EMEL JURULATIH"]).toLowerCase().trim() === e; });
}
function bolehUrusAtlet(emel) { return isAdmin(emel) || isJurulatihDilantik(emel); }
/* Padam atlet dari keseluruhan sistem — MASTER ADMIN sahaja */
function bolehPadamAtlet(emel) { return isAdmin(emel); }

function tambahAtlet(p) {
  if (!bolehUrusAtlet(p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih yang dilantik boleh menambah atlet baharu.");
  dapatSheet(SHEET_ATLET, HEADERS[SHEET_ATLET], "#00a3c4");
  var id = idBaharu("A", SHEET_ATLET);
  var urlGambar = p.gambar || "";
  if (p.gambarBase64) urlGambar = muatNaikGambar({ gambarBase64: p.gambarBase64, namaFail: id + "_" + String(p.nama).replace(/[^A-Za-z0-9]+/g, "_") }).url;
  ss().getSheetByName(SHEET_ATLET).appendRow([id, p.nama, p.noIc || "", p.jantina || "", p.kategori || "", p.sekolah || "", urlGambar, p.catatan || "", p.olehNama, nowStr(), Number(p.tinggi) || ""]);
  return { id: id, gambar: urlGambar };
}

/* Padam atlet — MASTER ADMIN sahaja.
   Rekod berkaitan (kehadiran, penyertaan, rekod acara, fail) turut dibersihkan. */
function padamAtlet(p) {
  if (!bolehPadamAtlet(p.olehEmel)) throw new Error("Hanya Master Admin boleh memadam atlet dari keseluruhan sistem.");
  if (!p || !p.id) throw new Error("ID atlet diperlukan.");
  var s = ss().getSheetByName(SHEET_ATLET);
  if (!s) throw new Error("Atlet tidak dijumpai.");
  var v = s.getDataRange().getValues(), jumpa = false;
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(p.id)) { s.deleteRow(i + 1); jumpa = true; break; }
  }
  if (!jumpa) throw new Error("Atlet tidak dijumpai.");

  /* Buang fail lampiran (Drive + baris sheet) */
  try {
    var sf = ss().getSheetByName(SHEET_FAIL);
    if (sf) {
      var vf = sf.getDataRange().getValues();
      for (var k = vf.length - 1; k >= 1; k--) {
        if (String(vf[k][1]) === String(p.id)) {
          try { DriveApp.getFileById(String(vf[k][6])).setTrashed(true); } catch (e) {}
          sf.deleteRow(k + 1);
        }
      }
    }
  } catch (e) {}

  /* Buang kehadiran & penyertaan (keluar dari SEMUA acara) */
  buangBarisMengikut(SHEET_KEHADIRAN, 2, p.id);
  buangBarisMengikut(SHEET_PENYERTAAN, 1, p.id);
  buangBarisMengikut(SHEET_BMI, 1, p.id);
  /* Buang semua rekod latihan atlet ini dalam setiap acara */
  baca(SHEET_ACARA).forEach(function (a) {
    if (a["ACARA"]) buangBarisMengikut(namaSheetRekod(a["ACARA"]), 3, p.id);
  });
  return { ok: true };
}

/* Keluarkan atlet dari SATU acara — jurulatih acara itu atau Master Admin */
function padamPenyertaan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh mengeluarkan atlet dari acara.");
  var s = ss().getSheetByName(SHEET_PENYERTAAN);
  if (!s) throw new Error("Penyertaan tidak dijumpai.");
  var v = s.getDataRange().getValues(), jumpa = false;
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(p.acara) && String(v[i][1]) === String(p.atletId)) { s.deleteRow(i + 1); jumpa = true; }
  }
  if (!jumpa) throw new Error("Atlet tiada dalam acara ini.");
  if (p.padamRekod) buangBarisMengikut(namaSheetRekod(p.acara), 3, p.atletId);
  return { ok: true };
}

/* Kemaskini satu rekod latihan — jurulatih acara itu atau Master Admin */
function kemaskiniRekod(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh mengedit rekod.");
  var s = sheetRekod(p.acara);
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      if (p.tarikh) v[i][1] = p.tarikh;
      if (p.keputusan !== undefined) v[i][7] = p.keputusan;
      if (p.nilai !== undefined) v[i][8] = Number(p.nilai) || "";
      if (p.catatan !== undefined) v[i][9] = p.catatan;
      v[i][10] = p.olehNama;
      v[i][11] = nowStr();
      s.getRange(i + 1, 1, 1, HEADER_REKOD.length).setValues([v[i].slice(0, HEADER_REKOD.length)]);
      var out = {};
      for (var j = 0; j < HEADER_REKOD.length; j++) out[HEADER_REKOD[j]] = v[i][j];
      return { ok: true, rekod: out };
    }
  }
  throw new Error("Rekod tidak dijumpai.");
}

/* Padam satu rekod latihan — jurulatih acara itu atau Master Admin */
function padamRekod(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh memadam rekod.");
  var s = sheetRekod(p.acara);
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(p.id)) { s.deleteRow(i + 1); return { ok: true }; }
  }
  throw new Error("Rekod tidak dijumpai.");
}

function buangBarisMengikut(namaSheet, indeksLajur, nilai) {
  try {
    var s = ss().getSheetByName(namaSheet);
    if (!s || s.getLastRow() < 2) return;
    var v = s.getDataRange().getValues();
    for (var i = v.length - 1; i >= 1; i--) {
      if (String(v[i][indeksLajur]) === String(nilai)) s.deleteRow(i + 1);
    }
  } catch (e) {}
}


function kemaskiniAtlet(p) {
  var s = ss().getSheetByName(SHEET_ATLET);
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      if (p.gambarBase64) {
        var g = muatNaikGambar({ gambarBase64: p.gambarBase64, namaFail: p.id + "_" + nowStr().replace(/[^0-9]/g, "") });
        p.gambar = g.url;
      }
      var medan = { nama: 1, noIc: 2, jantina: 3, kategori: 4, sekolah: 5, gambar: 6, catatan: 7, tinggi: 10 };
      // Kemaskini SEMUA medan yang dihantar (termasuk catatan kosong), supaya UI tidak perlu refresh.
      Object.keys(medan).forEach(function (k) { if (p[k] !== undefined && p[k] !== null) v[i][medan[k]] = p[k]; });
      v[i][8] = p.olehNama;
      v[i][9] = nowStr();
      s.getRange(i + 1, 1, 1, HEADERS[SHEET_ATLET].length).setValues([v[i].slice(0, HEADERS[SHEET_ATLET].length)]);
      var h = HEADERS[SHEET_ATLET], out = {};
      for (var j = 0; j < h.length; j++) out[h[j]] = v[i][j];
      return { ok: true, gambar: v[i][6], atlet: out };
    }
  }
  throw new Error("Atlet tidak dijumpai.");
}

/* (authorizeAll penuh ditakrifkan di atas — jangan tambah versi kedua,
   kerana takrifan kedua akan menimpa versi penuh tersebut.) */


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

/* ---------------- Auto Kehadiran ----------------
   Bila jurulatih mengisi rekod / menambah peserta, atlet terus dikira HADIR
   pada tarikh tersebut (tiada lagi sekatan "kehadiran belum ditanda"). */
function pastikanHadir_(atletId, nama, kategori, sekolah, tarikh, olehNama) {
  var t = tarikh || hariIni();
  dapatSheet(SHEET_KEHADIRAN, HEADERS[SHEET_KEHADIRAN], "#0f9d58");
  var s = ss().getSheetByName(SHEET_KEHADIRAN);
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (tarikhStr(v[i][1]) === t && String(v[i][2]) === String(atletId)) {
      if (String(v[i][6]).toUpperCase() === "HADIR") return { ok: true, sedia: true };
      s.getRange(i + 1, 7, 1, 4).setValues([["HADIR", "Auto — rekod latihan", olehNama || "", nowStr()]]);
      return { ok: true, auto: true, kemaskini: true };
    }
  }
  s.appendRow([idBaharu("K", SHEET_KEHADIRAN), t, atletId, nama || "", kategori || "", sekolah || "", "HADIR", "Auto — rekod latihan", olehNama || "", nowStr()]);
  return { ok: true, auto: true, baharu: true };
}

/* Migrasi ringan: pastikan sheet PENYERTAAN mempunyai kolum SUSUNAN. */
function pastikanKolumPenyertaan() {
  var s = dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var lc = Math.max(s.getLastColumn(), 1);
  var head = s.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).toUpperCase().trim(); });
  if (head.indexOf("SUSUNAN") >= 0) return s;
  var col = HEADERS[SHEET_PENYERTAAN].length;
  if (s.getMaxColumns() < col) s.insertColumnsAfter(s.getMaxColumns(), col - s.getMaxColumns());
  s.getRange(1, col).setValue("SUSUNAN").setFontWeight("bold").setFontColor("#ffffff").setBackground("#d81b60")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  s.setColumnWidth(col, 120);
  var lr = s.getLastRow();
  if (lr > 1) {
    var nilai = [];
    for (var i = 2; i <= lr; i++) nilai.push([i - 1]);
    s.getRange(2, col, nilai.length, 1).setValues(nilai);
  }
  return s;
}

/* Susunan peserta dalam SATU kategori bagi SATU acara (drag & drop).
   Jurulatih acara atau Master Admin sahaja. Hanya baris kategori itu diubah. */
function susunPeserta(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh menyusun peserta.");
  var kategori = String(p.kategori || "").toUpperCase().trim();
  var senarai = [];
  if (Object.prototype.toString.call(p.senarai) === "[object Array]") senarai = p.senarai.slice();
  else if (p.senarai) senarai = String(p.senarai).split("|");
  senarai = senarai.map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
  if (!senarai.length) throw new Error("Senarai susunan kosong.");

  var s = pastikanKolumPenyertaan();
  var col = HEADERS[SHEET_PENYERTAAN].length; /* kolum SUSUNAN */
  var v = s.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) !== String(p.acara)) continue;
    var kat = String(v[i][3] || "TANPA KATEGORI").toUpperCase().trim();
    if (kategori && kat !== kategori) continue; /* kategori lain tidak disentuh */
    var idx = senarai.indexOf(String(v[i][1]));
    if (idx < 0) continue;
    s.getRange(i + 1, col).setValue(idx + 1);
    n++;
  }
  return { ok: true, bil: n, kategori: kategori };
}

function lantikJurulatih(p) {
  /* Master Admin ATAU jurulatih acara itu boleh melantik jurulatih pembantu. */
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih acara ini boleh melantik jurulatih.");
  dapatSheet(SHEET_JURULATIH, HEADERS[SHEET_JURULATIH], "#ff6d00");
  var sedia = baca(SHEET_JURULATIH).filter(function (j) { return j["ACARA"] === p.acara; });
  if (sedia.length >= MAX_JURULATIH) throw new Error("Maksima " + MAX_JURULATIH + " jurulatih bagi setiap acara.");
  var e = String(p.emel).toLowerCase().trim();
  if (sedia.some(function (j) { return String(j["EMEL JURULATIH"]).toLowerCase() === e; })) throw new Error("Jurulatih ini sudah dilantik untuk acara tersebut.");
  ss().getSheetByName(SHEET_JURULATIH).appendRow([p.acara, e, p.nama, p.olehNama, nowStr()]);
  return { ok: true };
}

function buangJurulatih(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih acara ini boleh membuang jurulatih.");
  if (!isAdmin(p.olehEmel) && String(p.emel).toLowerCase() === String(p.olehEmel).toLowerCase()) throw new Error("Anda tidak boleh membuang diri sendiri sebagai jurulatih.");
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
  /* Auto: atlet yang dimasukkan ke acara terus dikira HADIR pada tarikh tersebut. */
  pastikanHadir_(p.atletId, p.nama, p.kategori, p.sekolah, p.tarikh || hariIni(), p.olehNama);
  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var ada = baca(SHEET_PENYERTAAN).some(function (x) { return x["ACARA"] === p.acara && String(x["ATLET ID"]) === String(p.atletId); });
  if (ada) throw new Error("Atlet sudah berada dalam acara ini.");
  pastikanKolumPenyertaan();
  var sp = ss().getSheetByName(SHEET_PENYERTAAN);
  var bilKat = baca(SHEET_PENYERTAAN).filter(function (x) {
    return x["ACARA"] === p.acara && String(x["KATEGORI"] || "").toUpperCase() === String(p.kategori || "").toUpperCase();
  }).length;
  sp.appendRow([p.acara, p.atletId, p.nama, p.kategori || "", p.sekolah || "", p.rekodPeribadi || "", p.olehNama, nowStr(), bilKat + 1]);
  sheetRekod(p.acara);
  return { ok: true };
}

/* Master Admin: tetapkan (ganti) senarai acara bagi SEORANG atlet sekaligus.
   Tidak memerlukan kehadiran hari ini — untuk kerja penyusunan pukal. */
function tetapkanAcaraAtlet(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh mengubah acara peserta secara pukal.");
  var atletId = String(p.atletId || "");
  if (!atletId) throw new Error("Atlet tidak dinyatakan.");

  var mahu = [];
  if (Object.prototype.toString.call(p.acara) === "[object Array]") mahu = p.acara.slice();
  else if (p.acara) mahu = String(p.acara).split("|");
  mahu = mahu.map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });

  var atlet = null, sen = baca(SHEET_ATLET);
  for (var k = 0; k < sen.length; k++) if (String(sen[k]["ID"]) === atletId) { atlet = sen[k]; break; }
  if (!atlet) throw new Error("Atlet tidak dijumpai.");

  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var s = ss().getSheetByName(SHEET_PENYERTAAN);
  var v = s.getDataRange().getValues();
  var sedia = {}, buang = 0;
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][1]) !== atletId) continue;
    var nm = String(v[i][0]);
    if (mahu.indexOf(nm) >= 0) { sedia[nm] = true; }
    else { s.deleteRow(i + 1); buang++; }
  }
  var tambah = 0;
  for (var j = 0; j < mahu.length; j++) {
    if (sedia[mahu[j]]) continue;
    s.appendRow([mahu[j], atletId, atlet["NAMA PENUH"], atlet["KATEGORI"] || "", atlet["SEKOLAH"] || "", "", p.olehNama, nowStr(), 9999]);
    sheetRekod(mahu[j]);
    tambah++;
  }
  return { ok: true, acara: mahu, tambah: tambah, buang: buang };
}

function kemaskiniRekodPeribadi(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh mengemaskini rekod peribadi.");
  dapatSheet(SHEET_PENYERTAAN, HEADERS[SHEET_PENYERTAAN], "#d81b60");
  var s = ss().getSheetByName(SHEET_PENYERTAAN);
  var v = s.getDataRange().getValues();
  var nilai = String(p.rekodPeribadi == null ? "" : p.rekodPeribadi).trim();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.acara) && String(v[i][1]) === String(p.atletId)) {
      s.getRange(i + 1, 6).setValue(nilai);
      return { ok: true, rekodPeribadi: nilai };
    }
  }
  throw new Error("Atlet tiada dalam acara ini.");
}

function simpanRekodLatihan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya jurulatih acara ini atau Master Admin boleh merekod.");
  var tarikh = p.tarikh || hariIni();

  /* Sokong acara berpasukan: p.atletIds = array id, p.namaList = array nama.
     Kekal serasi dengan panggilan lama yang hanya hantar p.atletId + p.nama. */
  var atletIds = [];
  if (Object.prototype.toString.call(p.atletIds) === "[object Array]") atletIds = p.atletIds.slice();
  else if (p.atletIds) atletIds = String(p.atletIds).split(",");
  else if (p.atletId) atletIds = [p.atletId];
  atletIds = atletIds.map(function (x) { return String(x).trim(); }).filter(function (x) { return x; });
  if (!atletIds.length) throw new Error("Tiada atlet dipilih untuk rekod ini.");

  var namaList = [];
  if (Object.prototype.toString.call(p.namaList) === "[object Array]") namaList = p.namaList.slice();
  else if (p.namaList) namaList = String(p.namaList).split("|");
  else if (p.nama) namaList = [p.nama];
  namaList = namaList.map(function (x) { return String(x).trim(); });
  while (namaList.length < atletIds.length) namaList.push(atletIds[namaList.length]);

  /* Auto kehadiran: setiap atlet dalam rekod ini terus ditanda HADIR. */
  var autoHadir = [];
  for (var i = 0; i < atletIds.length; i++) {
    var aid = atletIds[i];
    var rh = pastikanHadir_(aid, namaList[i] || "", p.kategori, p.sekolah, tarikh, p.olehNama);
    if (rh && rh.auto) autoHadir.push(aid);
  }

  var idJoin = atletIds.join(", ");
  var namaJoin = namaList.join(" + ");
  var s = sheetRekod(p.acara);
  var masa = Utilities.formatDate(new Date(), tz_(), "HH:mm:ss");
  var id = idBaharu("R", s.getName());
  var tms = nowStr();
  s.appendRow([id, tarikh, masa, idJoin, namaJoin, p.kategori || "", p.sekolah || "", p.keputusan, Number(p.nilai) || "", p.catatan || "", p.olehNama, tms]);
  return {
    ok: true,
    autoHadir: autoHadir,
    tarikhHadir: tarikh,
    rekod: {
      "ID": id, "TARIKH": tarikh, "MASA": masa, "ATLET ID": idJoin, "NAMA ATLET": namaJoin,
      "KATEGORI": p.kategori || "", "SEKOLAH": p.sekolah || "", "KEPUTUSAN": p.keputusan,
      "NILAI": Number(p.nilai) || "", "CATATAN": p.catatan || "", "DICATAT OLEH": p.olehNama, "TARIKH & MASA REKOD": tms
    }
  };
}

/* ---------------- Rekod Kejohanan (sasaran setiap acara + kategori) ----------------
   Hanya Master Admin dan jurulatih yang dilantik bagi acara berkenaan boleh
   menambah / mengemas kini / memadam rekod kejohanan.                        */
function simpanRekodKejohanan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih yang dilantik bagi acara ini boleh menetapkan rekod kejohanan.");
  var acara = String(p.acara || "").trim();
  var kategori = String(p.kategori || "").toUpperCase().trim();
  if (!acara) throw new Error("Acara diperlukan.");
  if (!kategori) throw new Error("Kategori diperlukan.");
  var nilai = Number(p.nilai);
  if (!(nilai > 0)) throw new Error("Nilai rekod tidak sah.");
  dapatSheet(SHEET_KEJOHANAN, HEADERS[SHEET_KEJOHANAN], "#b45309");
  var s = ss().getSheetByName(SHEET_KEJOHANAN);
  var tms = nowStr();
  var baris = [
    "", acara, kategori, p.namaKejohanan || "", p.tahun || "", p.pemegang || "",
    nilai, p.keputusan || "", p.catatan || "", p.olehNama || "", tms
  ];
  var peringkat = peringkatKejohanan(p.namaKejohanan);
  baris[3] = peringkat;
  var v = s.getDataRange().getValues();
  if (p.id) {
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][0]) === String(p.id)) {
        baris[0] = p.id;
        s.getRange(i + 1, 1, 1, baris.length).setValues([baris]);
        return { ok: true, dikemaskini: true, rekod: objKejohanan(baris) };
      }
    }
    throw new Error("Rekod kejohanan tidak dijumpai.");
  }
  /* Satu rekod bagi setiap acara + kategori + peringkat (MSSD / MSSK / NO.3 MSSK) — ganti jika sudah ada. */
  for (var j = 1; j < v.length; j++) {
    var pj = peringkatKejohanan(v[j][3]);
    if (String(v[j][1]) === acara && String(v[j][2]).toUpperCase() === kategori && pj === peringkat) {
      baris[0] = v[j][0];
      s.getRange(j + 1, 1, 1, baris.length).setValues([baris]);
      return { ok: true, dikemaskini: true, rekod: objKejohanan(baris) };
    }
  }
  baris[0] = idBaharu("RK", SHEET_KEJOHANAN);
  s.appendRow(baris);
  return { ok: true, rekod: objKejohanan(baris) };
}

function peringkatKejohanan(t) {
  var s = String(t || "MSSD").toUpperCase().replace(/\s+/g, " ").trim();
  if (/NO\.?\s*3/.test(s)) return "NO.3 MSSK";
  return s.indexOf("MSSK") >= 0 ? "MSSK" : "MSSD";
}

function objKejohanan(b) {
  var o = {};
  HEADERS[SHEET_KEJOHANAN].forEach(function (h, i) { o[h] = b[i]; });
  return o;
}

function padamRekodKejohanan(p) {
  if (!bolehRekod(p.acara, p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih yang dilantik bagi acara ini boleh memadam rekod kejohanan.");
  var s = ss().getSheetByName(SHEET_KEJOHANAN);
  if (!s) throw new Error("Rekod kejohanan tidak dijumpai.");
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(p.id)) { s.deleteRow(i + 1); return { ok: true }; }
  }
  throw new Error("Rekod kejohanan tidak dijumpai.");
}

function kolumAcara() {
  var s = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  var lc = Math.max(1, s.getLastColumn());
  var head = s.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x || "").toUpperCase().trim(); });
  var idx = {};
  head.forEach(function (h, i) { if (h) idx[h] = i + 1; });
  return { sheet: s, head: head, idx: idx };
}

function sisipKolumAcara(nama, isiFn) {
  /* Sisip kolum baharu (sebelum AKTIF jika ada) dan isikan nilai lalai. */
  var k = kolumAcara();
  if (k.idx[nama]) return;
  var s = k.sheet, lc = Math.max(1, s.getLastColumn());
  var iAktif = k.head.indexOf("AKTIF");
  var newCol;
  if (iAktif >= 0) { s.insertColumnBefore(iAktif + 1); newCol = iAktif + 1; }
  else { s.insertColumnAfter(lc); newCol = lc + 1; }
  s.getRange(1, newCol).setValue(nama);
  var lr = s.getLastRow();
  if (lr >= 2) {
    var namaAcara = s.getRange(2, 1, lr - 1, 1).getValues();
    var vals = namaAcara.map(function (r, i) { return [isiFn(String(r[0] || ""), i)]; });
    s.getRange(2, newCol, vals.length, 1).setValues(vals);
  }
}

function pastikanKolumAcara() {
  /* Migrasi ringan: pastikan sheet ACARA ada kolum MOD dan SUSUNAN (projek lama). */
  sisipKolumAcara("MOD", function (nm) {
    var relay = /(?:^|[^A-Z0-9])4\s*[xX]\s*\d+|MIX\s*RELAY|RELAY/i.test(nm);
    return relay ? "BERPASUKAN" : "INDIVIDU";
  });
  sisipKolumAcara("SUSUNAN", function (nm, i) { return i + 1; });
  /* Ikon acara (teks / gambar) — hanya Master Admin boleh ubah. */
  sisipKolumAcara("IKON", function (nm) { return String(nm || "").toUpperCase().replace(/[^A-Z0-9]/g, "").substring(0, 3); });
  sisipKolumAcara("IKON URL", function () { return ""; });
  sisipKolumAcara("IKON POS", function () { return ""; });
  /* Status pertandingan — acara yang digugurkan ditanda TIDAK. */
  sisipKolumAcara("DIPERTANDINGKAN", function () { return "YA"; });
}

/* Tetapkan beberapa nilai kolum bagi satu baris acara. petaNilai = {KOLUM: nilai} */
function tetapNilaiAcara_(nama, petaNilai) {
  pastikanKolumAcara();
  var k = kolumAcara(), s = k.sheet, lr = s.getLastRow();
  var cari = String(nama || "").toUpperCase().trim();
  if (!cari) throw new Error("Acara diperlukan.");
  for (var i = 2; i <= lr; i++) {
    if (String(s.getRange(i, 1).getValue()).toUpperCase().trim() === cari) {
      Object.keys(petaNilai).forEach(function (kol) {
        var c = k.idx[kol];
        if (c) s.getRange(i, c).setValue(petaNilai[kol]);
      });
      return { ok: true };
    }
  }
  throw new Error("Acara tidak dijumpai.");
}

/* Ikon acara — MASTER ADMIN sahaja (teks atau gambar yang dimuat naik). */
function tetapkanIkonAcara(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin & Sub Admin boleh mengubah ikon acara.");
  var nama = String(p.acara || "").trim();
  var teks = String(p.ikon || "").toUpperCase().trim().substring(0, 4);
  var url = normalGambar(p.ikonUrl || "");
  var pos = String(p.ikonPos || "").trim(); /* format: "x%|y%|zoom" */
  return tetapNilaiAcara_(nama, { "IKON": teks, "IKON URL": url, "IKON POS": pos });
}

/* Muat naik gambar ikon acara — MASTER ADMIN sahaja. */
function muatNaikIkonAcara(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin & Sub Admin boleh memuat naik ikon acara.");
  return muatNaikGambar({ gambarBase64: p.gambarBase64, namaFail: "IKON_" + String(p.acara || "ACARA").replace(/[^A-Za-z0-9]/g, "_") });
}

/* Tag acara sebagai TIDAK DIPERTANDINGKAN (digugurkan) — MASTER ADMIN sahaja. */
function tetapkanDipertandingkan(p) {
  var nama = String(p.acara || "").trim();
  if (!nama) throw new Error("Acara diperlukan.");
  if (!isAdmin(p.olehEmel)) {
    throw new Error("Hanya Master Admin & Sub Admin boleh menggugurkan acara.");
  }
  var nilai = String(p.nilai || "YA").toUpperCase() === "TIDAK" ? "TIDAK" : "YA";
  return tetapNilaiAcara_(nama, { "DIPERTANDINGKAN": nilai });
}

/* Gugurkan / kembalikan satu KATEGORI dalam satu ACARA.
   Hanya MASTER ADMIN & SUB ADMIN dibenarkan. */
function tetapkanKategoriGugur(p) {
  var acara = String(p.acara || "").toUpperCase().trim();
  var kategori = String(p.kategori || "").toUpperCase().trim();
  if (!acara || !kategori) throw new Error("Acara dan kategori diperlukan.");
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin & Sub Admin boleh menggugurkan kategori.");
  var status = String(p.status || "DIGUGURKAN").toUpperCase().trim() === "DIGUGURKAN" ? "DIGUGURKAN" : "AKTIF";
  var s = dapatSheet(SHEET_KAT_GUGUR, HEADERS[SHEET_KAT_GUGUR], "#b42318");
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).toUpperCase().trim() === acara && String(v[i][1]).toUpperCase().trim() === kategori) {
      s.deleteRow(i + 1);
    }
  }
  if (status === "DIGUGURKAN") {
    s.appendRow([acara, kategori, "DIGUGURKAN", p.olehNama || p.olehEmel || "", nowStr()]);
  }
  return { ok: true, acara: acara, kategori: kategori, status: status };
}

/* Susunan acara — MASTER ADMIN sahaja (drag & drop pada app). */
function susunAcara(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh menyusun kedudukan acara.");
  pastikanKolumAcara();
  var senarai = p.senarai || [];
  if (!senarai.length) throw new Error("Senarai susunan kosong.");
  var k = kolumAcara(), s = k.sheet, col = k.idx["SUSUNAN"];
  var lr = s.getLastRow();
  if (lr < 2 || !col) return { ok: true };
  var namaAcara = s.getRange(2, 1, lr - 1, 1).getValues();
  var peta = {};
  senarai.forEach(function (n, i) { peta[String(n).toUpperCase().trim()] = i + 1; });
  var vals = namaAcara.map(function (r, i) {
    var key = String(r[0] || "").toUpperCase().trim();
    return [peta[key] || (senarai.length + i + 1)];
  });
  s.getRange(2, col, vals.length, 1).setValues(vals);
  return { ok: true };
}

/* Padam acara sepenuhnya — MASTER ADMIN sahaja. */
function padamAcara(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh memadam acara.");
  var nama = String(p.acara || "").trim();
  if (!nama) throw new Error("Acara diperlukan.");
  var s = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  var v = s.getDataRange().getValues(), jumpa = false;
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).toUpperCase().trim() === nama.toUpperCase()) { s.deleteRow(i + 1); jumpa = true; }
  }
  if (!jumpa) throw new Error("Acara tidak dijumpai.");
  buangBarisMengikut(SHEET_JURULATIH, 0, nama);
  buangBarisMengikut(SHEET_PENYERTAAN, 0, nama);
  buangBarisMengikut(SHEET_KEJOHANAN, 1, nama);
  var sr = ss().getSheetByName(namaSheetRekod(nama));
  if (sr) ss().deleteSheet(sr);
  return { ok: true };
}

/* Jurulatih dilantik (mana-mana acara) atau Master Admin boleh menambah acara. */
function bolehTambahAcara(emel) { return isAdmin(emel) || isJurulatihDilantik(emel); }

function tambahAcara(p) {
  if (!bolehTambahAcara(p.olehEmel)) throw new Error("Hanya Master Admin atau jurulatih yang dilantik boleh menambah acara.");
  pastikanKolumAcara();
  var s = dapatSheet(SHEET_ACARA, HEADERS[SHEET_ACARA], "#6d28f9");
  var nama = String(p.acara).toUpperCase().trim();
  var sedia = baca(SHEET_ACARA);
  if (sedia.some(function (a) { return String(a["ACARA"]).toUpperCase() === nama; })) throw new Error("Acara sudah wujud.");
  var maxSusun = 0;
  sedia.forEach(function (a) { var n = Number(a["SUSUNAN"] || 0); if (n > maxSusun) maxSusun = n; });
  /* Tulis ikut kepala kolum supaya kekal betul walaupun ada kolum tambahan. */
  var k = kolumAcara();
  var nilai = {
    "ACARA": nama, "JENIS": p.jenis || "MASA", "UNIT": p.unit || "saat",
    "MOD": p.mod || "INDIVIDU", "SUSUNAN": maxSusun + 1,
    "IKON": nama.replace(/[^A-Z0-9]/g, "").substring(0, 3),
    "IKON URL": "", "IKON POS": "", "DIPERTANDINGKAN": "YA", "AKTIF": "YA"
  };
  var baris = [];
  for (var c = 0; c < k.head.length; c++) baris.push(nilai.hasOwnProperty(k.head[c]) ? nilai[k.head[c]] : "");
  s.appendRow(baris);
  sheetRekod(nama);
  /* Jurulatih yang menambah acara automatik dilantik sebagai jurulatih acara itu. */
  if (!isAdmin(p.olehEmel)) {
    dapatSheet(SHEET_JURULATIH, HEADERS[SHEET_JURULATIH], "#ff6d00");
    ss().getSheetByName(SHEET_JURULATIH).appendRow([nama, String(p.olehEmel).toLowerCase().trim(), p.olehNama || "", p.olehNama || "", nowStr()]);
  }
  return { ok: true };
}

/* ---------------- Router ---------------- */

/* ---------------- Pengurusan Sub Admin (Master Admin sahaja) ---------------- */
function tukarPerananGuru_(emel, peranan) {
  var e = String(emel || "").toLowerCase().trim();
  if (!e) throw new Error("Emel pengguna diperlukan.");
  if (e === ADMIN_EMEL) throw new Error("Peranan Master Admin tidak boleh diubah.");
  var s = ss().getSheetByName(SHEET_GURU);
  if (!s) throw new Error("Pengguna tidak dijumpai.");
  var v = s.getDataRange().getValues();
  var kolEmel = 2, kolPeranan = 7; /* ikut HEADERS[SHEET_GURU] */
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][kolEmel]).toLowerCase().trim() === e) {
      s.getRange(i + 1, kolPeranan + 1).setValue(peranan);
      return { ok: true, emel: e, peranan: peranan, nama: v[i][1] };
    }
  }
  throw new Error("Pengguna berdaftar tidak dijumpai.");
}

function lantikSubAdmin(p) {
  if (!isMasterAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh melantik Sub Admin.");
  return tukarPerananGuru_(p.emel, "SUBADMIN");
}

function buangSubAdmin(p) {
  if (!isMasterAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh menarik balik jawatan Sub Admin.");
  return tukarPerananGuru_(p.emel, "GURU");
}

/* Padam pengguna berdaftar. Master Admin & Sub Admin boleh, tetapi
   Master Admin aplikasi ini TIDAK BOLEH dipadam oleh sesiapa. */
function padamGuru(p) {
  if (!isAdmin(p.olehEmel)) throw new Error("Hanya Master Admin atau Sub Admin boleh memadam pengguna.");
  var e = String(p.emel || "").toLowerCase().trim();
  if (!e) throw new Error("Emel pengguna diperlukan.");
  if (e === ADMIN_EMEL) throw new Error("Master Admin aplikasi ini tidak boleh dipadam.");
  if (isSubAdmin(e) && !isMasterAdmin(p.olehEmel)) throw new Error("Hanya Master Admin boleh memadam Sub Admin.");
  if (e === String(p.olehEmel || "").toLowerCase().trim()) throw new Error("Anda tidak boleh memadam akaun sendiri.");
  var s = ss().getSheetByName(SHEET_GURU);
  if (!s) throw new Error("Pengguna tidak dijumpai.");
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][2]).toLowerCase().trim() === e) {
      s.deleteRow(i + 1);
      buangBarisMengikut(SHEET_JURULATIH, 1, e);
      return { ok: true };
    }
  }
  throw new Error("Pengguna tidak dijumpai.");
}

var TINDAKAN = {
  ping: function () { return { ok: true, masa: nowStr() }; },
  authorizeAll: function () { return { laporan: authorizeAll() }; },
  setup: function () { return { mesej: setupPangkalanData() }; },
  daftar: daftarGuru,
  login: login,
  data: semuaData,
  muatNaikGambar: muatNaikGambar,
  muatNaikFailAtlet: muatNaikFailAtlet,
  padamFailAtlet: padamFailAtlet,
  tambahAtlet: tambahAtlet,
  padamAtlet: padamAtlet,
  lantikSubAdmin: lantikSubAdmin,
  buangSubAdmin: buangSubAdmin,
  padamGuru: padamGuru,
  baikiUrlGambar: function () { return { mesej: baikiUrlGambar() }; },
  kemaskiniAtlet: kemaskiniAtlet,
  kehadiran: simpanKehadiran,
  lantikJurulatih: lantikJurulatih,
  buangJurulatih: buangJurulatih,
  tambahPenyertaan: tambahPenyertaan,
  padamPenyertaan: padamPenyertaan,
  kemaskiniRekodPeribadi: kemaskiniRekodPeribadi,
  tetapkanAcaraAtlet: tetapkanAcaraAtlet,
  rekod: simpanRekodLatihan,
  kemaskiniRekod: kemaskiniRekod,
  padamRekod: padamRekod,
  tambahAcara: tambahAcara,
  padamAcara: padamAcara,
  susunAcara: susunAcara,
  susunPeserta: susunPeserta,
  tetapkanIkonAcara: tetapkanIkonAcara,
  muatNaikIkonAcara: muatNaikIkonAcara,
  tetapkanDipertandingkan: tetapkanDipertandingkan,
  tetapkanKategoriGugur: tetapkanKategoriGugur,
  rekodKejohanan: simpanRekodKejohanan,
  padamRekodKejohanan: padamRekodKejohanan,
  tetapkanTinggi: tetapkanTinggi,
  bmi: simpanBmi,
  kemaskiniBmi: kemaskiniBmi,
  padamBmi: padamBmi
};

/* ================= BMI ATLET =================
   • TINGGI (CM) disimpan SEKALI dalam sheet ATLET (kolum "TINGGI (CM)").
   • BERAT boleh direkod bila-bila masa (setiap rekod = satu baris sheet BMI).
   • Master Admin, Sub Admin dan mana-mana jurulatih yang telah dilantik
     boleh mengubah tinggi & berat sekiranya berlaku kesalahan. */

function pastikanKolumTinggi() {
  var s = dapatSheet(SHEET_ATLET, HEADERS[SHEET_ATLET], "#00a3c4");
  var lc = Math.max(s.getLastColumn(), 1);
  var head = s.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x).toUpperCase().trim(); });
  if (head.indexOf("TINGGI (CM)") >= 0) return s;
  var col = lc + 1;
  if (s.getMaxColumns() < col) s.insertColumnsAfter(s.getMaxColumns(), col - s.getMaxColumns());
  s.getRange(1, col).setValue("TINGGI (CM)").setFontWeight("bold").setFontColor("#ffffff").setBackground("#00a3c4");
  return s;
}

function bolehUrusBmi(emel) { return isAdmin(emel) || isJurulatihDilantik(emel); }

function kiraBmi_(berat, tinggiCm) {
  var b = Number(berat) || 0, t = (Number(tinggiCm) || 0) / 100;
  if (!b || !t) return 0;
  return Math.round((b / (t * t)) * 10) / 10;
}
function statusBmi_(bmi) {
  var v = Number(bmi) || 0;
  if (!v) return "";
  if (v < 18.5) return "KURANG BERAT";
  if (v < 25) return "NORMAL";
  if (v < 30) return "BERLEBIHAN";
  return "OBES";
}

function tinggiAtlet_(atletId) {
  var s = ss().getSheetByName(SHEET_ATLET);
  if (!s || s.getLastRow() < 2) return { baris: 0, tinggi: 0, atlet: null };
  var v = s.getDataRange().getValues();
  var kol = v[0].map(function (x) { return String(x).toUpperCase().trim(); }).indexOf("TINGGI (CM)");
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(atletId)) {
      return { baris: i + 1, kol: kol + 1, tinggi: kol >= 0 ? (Number(v[i][kol]) || 0) : 0, atlet: v[i] };
    }
  }
  return { baris: 0, tinggi: 0, atlet: null };
}

/* Tetapkan / betulkan tinggi atlet */
function tetapkanTinggi(p) {
  if (!bolehUrusBmi(p.olehEmel)) throw new Error("Hanya Master Admin, Sub Admin atau jurulatih yang dilantik boleh menetapkan tinggi atlet.");
  pastikanKolumTinggi();
  var t = Number(p.tinggi) || 0;
  if (t < 80 || t > 250) throw new Error("Tinggi tidak munasabah (80–250 cm).");
  var info = tinggiAtlet_(p.atletId);
  if (!info.baris) throw new Error("Atlet tidak dijumpai.");
  if (info.tinggi && !p.paksa) throw new Error("Tinggi telah ditetapkan. Gunakan pilihan betulkan tinggi.");
  ss().getSheetByName(SHEET_ATLET).getRange(info.baris, info.kol).setValue(t);
  /* Kemaskini semula semua rekod BMI atlet ini supaya konsisten */
  kiraSemulaBmiAtlet_(p.atletId, t);
  return { ok: true, tinggi: t, bmi: baca(SHEET_BMI).filter(function (b) { return String(b["ATLET ID"]) === String(p.atletId); }) };
}

function kiraSemulaBmiAtlet_(atletId, tinggi) {
  var s = ss().getSheetByName(SHEET_BMI);
  if (!s || s.getLastRow() < 2) return;
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1]) === String(atletId)) {
      var bmi = kiraBmi_(v[i][7], tinggi);
      s.getRange(i + 1, 7, 1, 4).setValues([[tinggi, v[i][7], bmi, statusBmi_(bmi)]]);
    }
  }
}

/* Rekod berat baharu (bila-bila masa) */
function simpanBmi(p) {
  if (!bolehUrusBmi(p.olehEmel)) throw new Error("Hanya Master Admin, Sub Admin atau jurulatih yang dilantik boleh merekod BMI.");
  dapatSheet(SHEET_BMI, HEADERS[SHEET_BMI], "#0f766e");
  pastikanKolumTinggi();
  var info = tinggiAtlet_(p.atletId);
  if (!info.baris) throw new Error("Atlet tidak dijumpai.");
  var tinggi = Number(p.tinggi) || info.tinggi;
  if (!tinggi) throw new Error("Sila tetapkan tinggi atlet dahulu.");
  if (!info.tinggi && tinggi) ss().getSheetByName(SHEET_ATLET).getRange(info.baris, info.kol).setValue(tinggi);
  var berat = Number(p.berat) || 0;
  if (berat < 15 || berat > 250) throw new Error("Berat tidak munasabah (15–250 kg).");
  var bmi = kiraBmi_(berat, tinggi), st = statusBmi_(bmi);
  var id = idBaharu("B", SHEET_BMI);
  var tarikh = p.tarikh || hariIni();
  var baris = [id, p.atletId, p.nama || "", p.kategori || "", p.sekolah || "", tarikh, tinggi, berat, bmi, st, p.catatan || "", p.olehNama || "", nowStr()];
  ss().getSheetByName(SHEET_BMI).appendRow(baris);
  var out = {};
  for (var j = 0; j < HEADERS[SHEET_BMI].length; j++) out[HEADERS[SHEET_BMI][j]] = baris[j];
  return { ok: true, rekod: out, tinggi: tinggi };
}

/* Betulkan rekod BMI sedia ada */
function kemaskiniBmi(p) {
  if (!bolehUrusBmi(p.olehEmel)) throw new Error("Hanya Master Admin, Sub Admin atau jurulatih yang dilantik boleh mengubah rekod BMI.");
  var s = ss().getSheetByName(SHEET_BMI);
  if (!s) throw new Error("Rekod BMI tidak dijumpai.");
  var v = s.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(p.id)) {
      if (p.tarikh) v[i][5] = p.tarikh;
      if (p.tinggi) v[i][6] = Number(p.tinggi) || v[i][6];
      if (p.berat) v[i][7] = Number(p.berat) || v[i][7];
      if (p.catatan !== undefined) v[i][10] = p.catatan;
      v[i][8] = kiraBmi_(v[i][7], v[i][6]);
      v[i][9] = statusBmi_(v[i][8]);
      v[i][11] = p.olehNama || "";
      v[i][12] = nowStr();
      s.getRange(i + 1, 1, 1, HEADERS[SHEET_BMI].length).setValues([v[i].slice(0, HEADERS[SHEET_BMI].length)]);
      var out = {};
      for (var j = 0; j < HEADERS[SHEET_BMI].length; j++) out[HEADERS[SHEET_BMI][j]] = v[i][j];
      return { ok: true, rekod: out };
    }
  }
  throw new Error("Rekod BMI tidak dijumpai.");
}

function padamBmi(p) {
  if (!bolehUrusBmi(p.olehEmel)) throw new Error("Hanya Master Admin, Sub Admin atau jurulatih yang dilantik boleh memadam rekod BMI.");
  var s = ss().getSheetByName(SHEET_BMI);
  if (!s) throw new Error("Rekod BMI tidak dijumpai.");
  var v = s.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(p.id)) { s.deleteRow(i + 1); return { ok: true }; }
  }
  throw new Error("Rekod BMI tidak dijumpai.");
}

function balas(obj, callback) {
  var teks = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + teks + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(teks).setMimeType(ContentService.MimeType.JSON);
}

function kodRalat(mesej) {
  var m = String(mesej || "").toLowerCase();
  if (m.indexOf("tindakan tidak sah") >= 0) return "DB-100";
  if (m.indexOf("kata laluan") >= 0 || m.indexOf("emel") >= 0) return "DB-200";
  if (m.indexOf("hanya") >= 0 || m.indexOf("maksima") >= 0) return "DB-300";
  if (m.indexOf("kehadiran") >= 0) return "DB-400";
  if (m.indexOf("tidak dijumpai") >= 0 || m.indexOf("sudah") >= 0) return "DB-500";
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
  } catch (x) {}
}

/* Aksi ringan tidak perlu Script Lock — bolehkan parallelism supaya kehadiran laju */
var TANPA_LOCK = { ping: 1, data: 1, login: 1 };


/* ============ CACHE PANTAS (Script Cache) ============
   Balasan "data" disimpan dalam cache supaya pemuatan jauh lebih laju.
   Cache dibatalkan secara automatik setiap kali ada penulisan data. */
var CACHE_KEY = "AT_DATA_V1";
var CACHE_TTL = 15; /* saat — pendek supaya perubahan Sheet cepat sampai ke telefon */

function cacheBaca_() {
  try {
    var c = CacheService.getScriptCache();
    var meta = c.get(CACHE_KEY);
    if (!meta) return null;
    var n = Number(meta), keys = [];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY + "_" + i);
    var m = c.getAll(keys), s = "";
    for (i = 0; i < n; i++) { if (m[CACHE_KEY + "_" + i] == null) return null; s += m[CACHE_KEY + "_" + i]; }
    return s;
  } catch (e) { return null; }
}
function cacheTulis_(s) {
  try {
    var c = CacheService.getScriptCache(), size = 90000, n = Math.ceil(s.length / size);
    if (n > 40) return; /* terlalu besar untuk cache */
    var o = {};
    for (var i = 0; i < n; i++) o[CACHE_KEY + "_" + i] = s.substr(i * size, size);
    c.putAll(o, CACHE_TTL);
    c.put(CACHE_KEY, String(n), CACHE_TTL);
  } catch (e) {}
}
function cacheBatal_() { try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) {} }

function balasMentah_(json, callback) {
  if (callback) return ContentService.createTextOutput(callback + "(" + json + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function proses(payload, callback) {
  var mula = new Date().getTime();
  try {
    var fn = TINDAKAN[payload.action];
    if (!fn) throw new Error("Tindakan tidak sah: " + payload.action);

    /* Bacaan penuh: guna cache supaya balasan hampir serta-merta */
    if (payload.action === "data" && !payload.paksa) {
      var raw = cacheBaca_();
      if (raw) return balasMentah_('{"ok":true,"cache":1,"data":' + raw + '}', callback);
      raw = JSON.stringify(fn(payload));
      cacheTulis_(raw);
      return balasMentah_('{"ok":true,"data":' + raw + '}', callback);
    }

    if (TANPA_LOCK[payload.action]) {
      return balas({ ok: true, data: fn(payload), ms: new Date().getTime() - mula }, callback);
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var hasil = fn(payload);
      cacheBatal_(); /* data berubah — batalkan cache supaya semua peranti dapat data terkini */
      return balas({ ok: true, data: hasil, ms: new Date().getTime() - mula }, callback);
    }
    finally { lock.releaseLock(); }
  } catch (err) {
    logRalat(payload, err);
    return balas({
      ok: false,
      kod: kodRalat(err && err.message),
      error: String((err && err.message) || err),
      aksi: (payload && payload.action) || "-",
      masa: nowStr(),
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
