/**
 * LATIHAN PUSAT OLAHRAGA - Backend Google Apps Script
 * ---------------------------------------------------
 * Cara pasang:
 *  1. Buka https://script.google.com > New Project
 *  2. Ganti isi Code.gs dengan fail ini
 *  3. Klik Run > setup (benarkan akses)
 *  4. Deploy > New deployment > Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Salin Web App URL, letakkan dalam index.html (const API_URL = ...)
 *
 * Struktur Google Sheet (dibuat automatik):
 *   - GURU          : nama, email, password, jawatan, sekolah, telefon, role, daftar_pada
 *   - ALTELIT       : nama, ic, jantina, kategori, sekolah, foto, rekod_peribadi(JSON), didaftar_oleh, pada
 *   - KEHADIRAN     : tarikh, nama_altelit, catat_oleh, masa
 *   - ACARA_JURULATIH: acara, senarai_jurulatih(JSON, max 10)
 *   - Satu sheet setiap acara (contoh: "100 M", "LOMPAT JAUH"):
 *       tarikh, masa, nama_altelit, kategori, sekolah, rekod, catat_oleh
 */

var SHEET_ID = ''; // kosongkan untuk cipta baharu automatik pada run pertama
var MASTER_PW = '101010';
var ACARA_LIST = ['100 M','200 M','400 M','800 M','1500 M','LOMPAT JAUH','LOMPAT TINGGI','LONTAR PELURU','LEMPAR CAKERA','REJAM LEMBING'];

function setup() {
  var ss = getSS();
  ensureSheet(ss, 'GURU', ['nama','email','password','jawatan','sekolah','telefon','role','daftar_pada']);
  ensureSheet(ss, 'ALTELIT', ['nama','ic','jantina','kategori','sekolah','foto','rekod_peribadi','didaftar_oleh','pada']);
  ensureSheet(ss, 'KEHADIRAN', ['tarikh','nama_altelit','catat_oleh','masa']);
  ensureSheet(ss, 'ACARA_JURULATIH', ['acara','senarai_jurulatih']);
  ACARA_LIST.forEach(function(a){
    ensureSheet(ss, a, ['tarikh','masa','nama_altelit','kategori','sekolah','rekod','catat_oleh']);
  });
  Logger.log('Setup selesai. Sheet ID: ' + ss.getId());
  return ss.getId();
}

function getSS() {
  var props = PropertiesService.getScriptProperties();
  var id = SHEET_ID || props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e){}
  }
  var ss = SpreadsheetApp.create('LATIHAN PUSAT OLAHRAGA - Pengkalan Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function ensureSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var out = { ok:false };
  try {
    var params = (e && e.parameter) || {};
    if (e && e.postData && e.postData.contents) {
      try { var body = JSON.parse(e.postData.contents); for (var k in body) params[k]=body[k]; } catch(x){}
    }
    var action = params.action || 'ping';
    out = ({
      ping: pingFn, register: registerFn, login: loginFn,
      listAthletes: listAthletesFn, addAthlete: addAthleteFn,
      markAttendance: markAttendanceFn, listAttendanceToday: listAttendanceTodayFn,
      saveRecord: saveRecordFn, listRecords: listRecordsFn, listRecordsByAthlete: listRecordsByAthleteFn,
      listCoaches: listCoachesFn, setCoaches: setCoachesFn,
      dashboard: dashboardFn, bulkSeed: bulkSeedFn
    }[action] || pingFn)(params);
    out.ok = true;
  } catch(err) {
    out = { ok:false, error: String(err) };
  }
  var callback = (e && e.parameter && e.parameter.callback);
  var json = JSON.stringify(out);
  if (callback) {
    return ContentService.createTextOutput(callback+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function pingFn(){ return { message:'LATIHAN PUSAT API aktif', time:new Date().toISOString() }; }

/* ---- AUTH ---- */
function registerFn(p) {
  var sh = ensureSheet(getSS(),'GURU',['nama','email','password','jawatan','sekolah','telefon','role','daftar_pada']);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){ if ((data[i][1]||'').toLowerCase()===String(p.email||'').toLowerCase()) throw 'Email sudah berdaftar'; }
  var role = (String(p.masterPassword||'') === MASTER_PW) ? 'MASTER_ADMIN' : 'JURULATIH';
  sh.appendRow([p.nama,p.email,p.password,p.jawatan,p.sekolah,p.telefon,role,new Date()]);
  return { user:{ nama:p.nama,email:p.email,jawatan:p.jawatan,sekolah:p.sekolah,telefon:p.telefon,role:role } };
}
function loginFn(p){
  var sh = getSS().getSheetByName('GURU'); if(!sh) throw 'Tiada rekod guru';
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){
    if ((data[i][1]||'').toLowerCase()===String(p.email||'').toLowerCase() && String(data[i][2])===String(p.password)){
      return { user:{ nama:data[i][0],email:data[i][1],jawatan:data[i][3],sekolah:data[i][4],telefon:data[i][5],role:data[i][6] } };
    }
  }
  throw 'Email atau kata laluan salah';
}

/* ---- ALTELIT ---- */
function listAthletesFn(){
  var sh = getSS().getSheetByName('ALTELIT'); if(!sh) return { athletes:[] };
  var data = sh.getDataRange().getValues(); var out=[];
  for (var i=1;i<data.length;i++){
    var rekod={}; try{ rekod=JSON.parse(data[i][6]||'{}'); }catch(x){}
    out.push({ nama:data[i][0], ic:data[i][1], jantina:data[i][2], kategori:data[i][3], sekolah:data[i][4], foto:data[i][5], rekod:rekod });
  }
  return { athletes: out };
}
function addAthleteFn(p){
  if (p.role !== 'MASTER_ADMIN') throw 'Hanya Master Admin boleh tambah altelit';
  var sh = ensureSheet(getSS(),'ALTELIT',['nama','ic','jantina','kategori','sekolah','foto','rekod_peribadi','didaftar_oleh','pada']);
  sh.appendRow([p.nama,p.ic||'',p.jantina||'L',p.kategori||'',p.sekolah||'',p.foto||'', JSON.stringify(p.rekod||{}), p.by||'', new Date()]);
  return { added:true };
}

/* ---- KEHADIRAN ---- */
function markAttendanceFn(p){
  var sh = ensureSheet(getSS(),'KEHADIRAN',['tarikh','nama_altelit','catat_oleh','masa']);
  var tarikh = p.tarikh || Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'yyyy-MM-dd');
  // Buang rekod hari itu untuk altelit tersebut (elak dup)
  var data = sh.getDataRange().getValues();
  for (var i=data.length-1;i>=1;i--){
    var t = data[i][0]; var ts = (t instanceof Date) ? Utilities.formatDate(t,Session.getScriptTimeZone(),'yyyy-MM-dd') : String(t);
    if (ts===tarikh && data[i][1]===p.nama) sh.deleteRow(i+1);
  }
  sh.appendRow([tarikh, p.nama, p.by||'', new Date()]);
  return { marked:true };
}
function listAttendanceTodayFn(p){
  var sh = getSS().getSheetByName('KEHADIRAN'); if(!sh) return { hadir:[] };
  var tarikh = p.tarikh || Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'yyyy-MM-dd');
  var data = sh.getDataRange().getValues(); var out=[];
  for (var i=1;i<data.length;i++){
    var t = data[i][0]; var ts = (t instanceof Date) ? Utilities.formatDate(t,Session.getScriptTimeZone(),'yyyy-MM-dd') : String(t);
    if (ts===tarikh) out.push({ nama:data[i][1], by:data[i][2], masa:data[i][3] });
  }
  return { hadir: out, tarikh: tarikh };
}

/* ---- REKOD LATIHAN ---- */
function saveRecordFn(p){
  // pastikan altelit hadir hari ini
  var tarikh = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  var kh = getSS().getSheetByName('KEHADIRAN'); var hadir=false;
  if (kh){ var d=kh.getDataRange().getValues();
    for (var i=1;i<d.length;i++){ var t=d[i][0]; var ts=(t instanceof Date)?Utilities.formatDate(t,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(t);
      if (ts===tarikh && d[i][1]===p.nama){ hadir=true; break; } } }
  if (!hadir) throw 'Altelit belum ditandai hadir hari ini';

  // pastikan jurulatih dibenarkan untuk acara ini (jika ada senarai)
  var cs = getSS().getSheetByName('ACARA_JURULATIH');
  if (cs){
    var cd = cs.getDataRange().getValues();
    for (var j=1;j<cd.length;j++){
      if (cd[j][0]===p.acara){
        var list=[]; try{ list=JSON.parse(cd[j][1]||'[]'); }catch(x){}
        if (list.length>0 && p.byRole!=='MASTER_ADMIN' && list.indexOf(p.byEmail)<0) throw 'Anda bukan jurulatih untuk acara ini';
        break;
      }
    }
  }
  var sh = ensureSheet(getSS(),p.acara,['tarikh','masa','nama_altelit','kategori','sekolah','rekod','catat_oleh']);
  sh.appendRow([tarikh, new Date(), p.nama, p.kategori||'', p.sekolah||'', p.rekod, p.by||'']);
  return { saved:true };
}
function listRecordsFn(p){
  var sh = getSS().getSheetByName(p.acara); if(!sh) return { rekod:[] };
  var d = sh.getDataRange().getValues(); var out=[];
  for (var i=1;i<d.length;i++) out.push({ tarikh:formatDT(d[i][0]), masa:formatDT(d[i][1]), nama:d[i][2], kategori:d[i][3], sekolah:d[i][4], rekod:d[i][5], by:d[i][6] });
  return { rekod: out.reverse() };
}
function listRecordsByAthleteFn(p){
  var out={}; ACARA_LIST.forEach(function(a){
    var sh=getSS().getSheetByName(a); if(!sh) return;
    var d=sh.getDataRange().getValues(); var arr=[];
    for (var i=1;i<d.length;i++){ if (d[i][2]===p.nama) arr.push({ tarikh:formatDT(d[i][0]), rekod:d[i][5], by:d[i][6] }); }
    if (arr.length) out[a]=arr;
  });
  return { rekod: out };
}
function formatDT(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
  return String(v||'');
}

/* ---- JURULATIH ACARA ---- */
function listCoachesFn(){
  var sh = ensureSheet(getSS(),'ACARA_JURULATIH',['acara','senarai_jurulatih']);
  var d = sh.getDataRange().getValues(); var m={};
  for (var i=1;i<d.length;i++){ try{ m[d[i][0]]=JSON.parse(d[i][1]||'[]'); }catch(x){ m[d[i][0]]=[]; } }
  return { coaches:m };
}
function setCoachesFn(p){
  if (p.role!=='MASTER_ADMIN') throw 'Hanya Master Admin';
  var arr = (p.list||[]).slice(0,10);
  var sh = ensureSheet(getSS(),'ACARA_JURULATIH',['acara','senarai_jurulatih']);
  var d = sh.getDataRange().getValues(); var found=false;
  for (var i=1;i<d.length;i++){ if (d[i][0]===p.acara){ sh.getRange(i+1,2).setValue(JSON.stringify(arr)); found=true; break; } }
  if (!found) sh.appendRow([p.acara, JSON.stringify(arr)]);
  return { saved:true };
}

/* ---- DASHBOARD ---- */
function dashboardFn(){
  var out={};
  ACARA_LIST.forEach(function(a){
    var sh=getSS().getSheetByName(a); if(!sh) return;
    var d=sh.getDataRange().getValues(); var agg={};
    for (var i=1;i<d.length;i++){
      var nama=d[i][2]; var v=parseRecord(d[i][5]); if (v===null) continue;
      if (!agg[nama]) agg[nama]={ sekolah:d[i][4], kategori:d[i][3], sum:0, n:0, best:null };
      agg[nama].sum+=v; agg[nama].n++;
      agg[nama].best = (agg[nama].best===null) ? v : (isTimeEvent(a) ? Math.min(agg[nama].best,v) : Math.max(agg[nama].best,v));
    }
    var rows=Object.keys(agg).map(function(n){ var x=agg[n]; return { nama:n, sekolah:x.sekolah, kategori:x.kategori, purata:(x.sum/x.n), best:x.best, count:x.n }; });
    rows.sort(function(a,b){ return isTimeEvent(a.acara)? a.purata-b.purata : b.purata-a.purata; });
    // sort based on event
    var time = isTimeEvent(a);
    rows.sort(function(x,y){ return time ? x.purata-y.purata : y.purata-x.purata; });
    out[a]=rows;
  });
  return { ranking: out };
}
function isTimeEvent(a){ return /M$/.test(a); }
function parseRecord(s){
  if (s===''||s==null) return null;
  var m = String(s).match(/([0-9]+(?:\.[0-9]+)?)/);
  return m ? parseFloat(m[1]) : null;
}

/* ---- BULK SEED (jalankan sekali dari index) ---- */
function bulkSeedFn(p){
  var sh = ensureSheet(getSS(),'ALTELIT',['nama','ic','jantina','kategori','sekolah','foto','rekod_peribadi','didaftar_oleh','pada']);
  if (sh.getLastRow()>1) return { skipped:true };
  var arr = p.athletes || [];
  var rows = arr.map(function(a){ return [a.nama,a.ic||'',a.jantina||'L',a.kategori||'',a.sekolah||'','', JSON.stringify(a.acara||{}),'SEED',new Date()]; });
  if (rows.length) sh.getRange(2,1,rows.length,9).setValues(rows);
  return { seeded: rows.length };
}
