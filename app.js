/* =====================================================================
   ENTRI DATA -> GOOGLE SPREADSHEET        (HTML + CSS + JavaScript saja)

   Cara kerja singkat
   - Aplikasi ini berjalan sepenuhnya di browser. Data ditulis ke Google Sheets
     lewat Google Sheets API memakai akun Google orang yang masuk (OAuth).
   - Akun pengguna (username + hash password) dan pengaturan kolom disimpan di
     tab tersembunyi "_pengaturan" pada "spreadsheet utama", sehingga HP dan laptop
     memakai akun dan pengaturan yang sama.
   - Angka 0 di depan dijaga dengan petik 1 (') di depan nilai, lalu dikirim
     dengan valueInputOption=USER_ENTERED supaya Sheets menyimpannya sebagai teks.

   Catatan keamanan: login username/password di sini hanyalah pembatas tampilan.
   Pengaman yang sebenarnya adalah siapa saja yang kamu beri akses Editor ke spreadsheet.
   ===================================================================== */
'use strict';

/* ---------- PENGATURAN BAWAAN ---------- */
// Opsional: isi Client ID di sini supaya semua perangkat tidak perlu mengetiknya.
const GOOGLE_CLIENT_ID = '';

// Akun admin pertama. Yang disimpan hanya hash-nya; password awal diberikan terpisah.
// Admin WAJIB mengganti password saat masuk pertama kali.
const ADMIN_USERNAME = 'admin';
const ADMIN_SALT = '2+zY34ThKazWZ/AKnl8+HQ==';
const ADMIN_HASH = 'MZtYh8sVhxGWDpT4hyIbv58tKcz84/ARvHoLiyRi5cM=';

const ITERASI = 120000;
const ALFABET_PW = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
const TIPE_VALID = ['teks', 'angka', 'kode', 'tanggal', 'catatan', 'pilihan'];

/* ==== LOGIKA MURNI (tanpa DOM) ==== */
// Angka: terima koma/titik desimal, tolak selain angka. Hasil memakai titik, atau null bila tidak valid.
function normAngka(raw){
  const s = String(raw).replace(/\s/g, '').replace(',', '.');
  return /^-?\d+(\.\d+)?$/.test(s) ? s : null;
}
// Petik 1 di depan = paksa Sheets/Excel menyimpan sebagai TEKS.
function amankanTeks(v){
  if (v === '' || v.charAt(0) === "'") return v;
  if (/^0\d/.test(v)) return "'" + v;          // 0812…, 007 -> angka 0 di depan tetap ada
  if (/^\d{12,}$/.test(v)) return "'" + v;     // NIK / rekening panjang, cegah jadi 3,2E+15
  if (/^[=+@]/.test(v)) return "'" + v;        // cegah dibaca rumus; "+62…" tidak kehilangan plus
  if (v.charAt(0) === '-' && isNaN(Number(v.replace(',', '.')))) return "'" + v;
  return v;
}
// Mengembalikan nilai siap kirim (teks atau angka), atau null bila angka tidak valid.
function siapkan(tipe, v){
  if (v === '') return '';
  if (tipe === 'angka'){
    const n = normAngka(v);
    if (n === null) return null;
    if (/^-?0\d/.test(n) || n.replace(/[-.]/g, '').length > 15) return "'" + n;  // "007" / terlalu panjang -> teks
    return Number(n);                                                           // angka sungguhan, bebas locale
  }
  if (tipe === 'kode' && /^\d+$/.test(v)) return "'" + v;   // semua digit -> selalu teks
  return amankanTeks(v);
}
function colLetter(n){
  let s = '';
  while (n > 0){ const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function kutip(nama){ return "'" + String(nama).replace(/'/g, "''") + "'"; }
function rng(nama, a1){ return encodeURIComponent(kutip(nama) + '!' + a1); }
function ambilId(url){
  const s = String(url || '').trim();
  const m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;
  return '';
}
function urlDariId(id){ return 'https://docs.google.com/spreadsheets/d/' + id + '/edit'; }
function cekNamaSheet(n){
  if (n.length > 100) return 'Nama sheet maksimal 100 karakter.';
  if (/[\[\]*?:\/\\]/.test(n)) return 'Nama sheet tidak boleh memuat [ ] * ? : / \\';
  if (n.toLowerCase() === '_pengaturan') return 'Nama “_pengaturan” dipakai aplikasi. Pilih nama lain.';
  return '';
}
function pad2(n){ return n < 10 ? '0' + n : '' + n; }
function formatWaktu(d){
  return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) + ' ' +
         pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
}
function cariKolom(judul, nama){
  const n = String(nama).trim().toLowerCase();
  for (let i = 0; i < judul.length; i++) if (String(judul[i]).trim().toLowerCase() === n) return i;
  return -1;
}
// Tambahkan nama kolom yang belum ada di judul (di kanan). Mengembalikan judul baru + daftar kolom baru.
function gabungHeader(judul, items){
  const h = judul.slice(), baru = [];
  items.forEach(function(it){ if (cariKolom(h, it.nama) === -1){ h.push(it.nama); baru.push(it.nama); } });
  return {judul: h, baru: baru, awal: judul.length};
}
function susunBaris(judul, items){
  const baris = judul.map(function(){ return ''; });
  items.forEach(function(it){ const i = cariKolom(judul, it.nama); if (i > -1) baris[i] = it.nilai; });
  return baris;
}
// Kolom yang berformat "Teks biasa" sudah menjaga angka 0; petik 1 justru ikut tampil -> buang.
function lepasPetik(baris, kolomTeks){
  Object.keys(kolomTeks).forEach(function(i){
    const v = baris[i];
    if (typeof v === 'string' && v.charAt(0) === "'") baris[i] = v.slice(1);
  });
  return baris;
}
function bytesKeB64(b){ let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]); return btoa(s); }
function b64KeBytes(s){ const r = atob(s), b = new Uint8Array(r.length); for (let i = 0; i < r.length; i++) b[i] = r.charCodeAt(i); return b; }
function adaCrypto(){ return typeof crypto !== 'undefined' && !!crypto.subtle && !!crypto.getRandomValues; }
async function hashPw(pw, saltB64){
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({name: 'PBKDF2', hash: 'SHA-256', salt: b64KeBytes(saltB64), iterations: ITERASI}, key, 256);
  return bytesKeB64(new Uint8Array(bits));
}
function samaPersis(a, b){
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function saltBaru(){ return bytesKeB64(crypto.getRandomValues(new Uint8Array(16))); }
function passwordAcak(){
  let out = '';
  while (out.length < 12){
    const b = crypto.getRandomValues(new Uint8Array(24));
    for (let i = 0; i < b.length && out.length < 12; i++) if (b[i] < 224) out += ALFABET_PW.charAt(b[i] % 56);
  }
  return out.slice(0, 4) + '-' + out.slice(4, 8) + '-' + out.slice(8, 12);
}
function buatTokenSetup(clientId, sheetId){
  return btoa(JSON.stringify({c: clientId, s: sheetId})).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function bacaTokenSetup(t){
  try{
    let b = String(t).replace(/-/g, '+').replace(/_/g, '/');
    while (b.length % 4) b += '=';
    const j = JSON.parse(atob(b));
    if (typeof j.c === 'string' && /^[\w.-]+$/.test(j.c) && typeof j.s === 'string' && /^[\w-]+$/.test(j.s)) return {clientId: j.c, sheetId: j.s};
  }catch(e){}
  return null;
}
function konfigBawaan(){
  return {v: 1, target: '', sheet: 'Data', waktu: true, catatUser: true, kolom: [
    {nama: 'Nama', tipe: 'teks', opsi: ''},
    {nama: 'No. HP', tipe: 'kode', opsi: ''},
    {nama: 'Keterangan', tipe: 'catatan', opsi: ''}
  ]};
}
function bersihKonfig(k){
  if (!k || typeof k !== 'object' || !Array.isArray(k.kolom)) return null;
  return {
    v: 1,
    target: typeof k.target === 'string' ? k.target : '',
    sheet: (typeof k.sheet === 'string' && k.sheet.trim()) ? k.sheet.trim() : 'Data',
    waktu: !!k.waktu,
    catatUser: k.catatUser !== false,
    kolom: k.kolom.filter(function(x){ return x && typeof x.nama === 'string' && x.nama.trim(); }).map(function(x){
      return {nama: x.nama.trim(), tipe: TIPE_VALID.indexOf(x.tipe) > -1 ? x.tipe : 'teks', opsi: typeof x.opsi === 'string' ? x.opsi : ''};
    })
  };
}
function adminBawaan(){
  return {u: ADMIN_USERNAME, salt: ADMIN_SALT, hash: ADMIN_HASH, role: 'admin', aktif: true, ganti: true, dibuat: new Date().toISOString()};
}
function jsonUser(u){
  return JSON.stringify({u: u.u, salt: u.salt, hash: u.hash, role: u.role, aktif: !!u.aktif, ganti: !!u.ganti, dibuat: u.dibuat});
}
function cekUsername(u){ return /^[a-z0-9._-]{3,30}$/.test(u); }
/* ==== AKHIR LOGIKA MURNI ==== */

(function(){

const SCOPE = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email';
const TAB_PUSAT = '_pengaturan';
const KOL_WAKTU = 'Waktu Input';
const KOL_USER = 'Diinput Oleh';
const MAKS_GAGAL = 5;
const KUNCI_MS = 15 * 60 * 1000;
const TIPE = [
  {v: 'teks',    l: 'Teks pendek'},
  {v: 'angka',   l: 'Angka (bisa dihitung)'},
  {v: 'kode',    l: 'Nomor / kode (HP, NIK, rekening)'},
  {v: 'tanggal', l: 'Tanggal'},
  {v: 'catatan', l: 'Teks panjang'},
  {v: 'pilihan', l: 'Pilihan (dropdown)'}
];

const $ = function(id){ return document.getElementById(id); };
let uidN = 0;
const uid = function(){ return ++uidN; };

/* ---------- penyimpanan di browser (dengan cadangan di memori) ---------- */
const mem = {local: {}, sess: {}};
function simpanKe(jenis, k, v){
  try{
    const st = jenis === 'local' ? window.localStorage : window.sessionStorage;
    if (v === null) st.removeItem(k); else st.setItem(k, JSON.stringify(v));
  }catch(e){}
  mem[jenis][k] = v;
}
function bacaDari(jenis, k){
  try{
    const st = jenis === 'local' ? window.localStorage : window.sessionStorage;
    const raw = st.getItem(k);
    if (raw !== null) return JSON.parse(raw);
  }catch(e){}
  return mem[jenis][k] === undefined ? null : mem[jenis][k];
}

/* ---------- state ---------- */
const S = {clientId: '', homeUrl: '', homeId: '', judul: '', judulTarget: '', email: '',
           users: [], konfig: konfigBawaan(), konfigRow: 2, konfigAda: false};
let C = {urlInput: '', sheet: 'Data', waktu: true, catatUser: true};   // salinan kerja pengaturan (admin)
let K = [];                                                            // salinan kerja kolom
let ME = null;                                                         // {username, role, ganti, hp}
let tok = {t: '', exp: 0};
let tokenJanji = null;
let dirty = false, sedangKirim = false, jumlahSesi = 0;
const cacheFmt = {};

/* ---------- pembantu DOM ---------- */
function el(tag, attrs){
  const e = document.createElement(tag);
  const a = attrs || {};
  Object.keys(a).forEach(function(k){
    const v = a[k];
    if (v === false || v === null || v === undefined) return;
    if (k === 'value' || k === 'selected' || k === 'checked') e[k] = v;
    else e.setAttribute(k, v === true ? '' : v);
  });
  for (let i = 2; i < arguments.length; i++) e.append(arguments[i]);
  return e;
}
let toastT;
function toast(msg, err){
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show' + (err ? ' err' : '');
  clearTimeout(toastT);
  toastT = setTimeout(function(){ t.className = 'toast'; }, err ? 6000 : 2600);
}
function sibuk(btn, ya, teks){
  if (!btn.dataset.label) btn.dataset.label = btn.textContent;
  btn.disabled = ya;
  btn.textContent = ya ? teks : btn.dataset.label;
}
// Tombol dua langkah (pengganti confirm())
function duaKetuk(btn, label, fn){
  btn.addEventListener('click', function(){
    if (btn.dataset.minta === '1'){
      clearTimeout(btn._t); btn.dataset.minta = ''; btn.textContent = label; fn(); return;
    }
    btn.dataset.minta = '1'; btn.textContent = 'Yakin? Ketuk lagi';
    btn._t = setTimeout(function(){ btn.dataset.minta = ''; btn.textContent = label; }, 3500);
  });
}
function tampil(v){
  ['muat', 'setup', 'login', 'app'].forEach(function(x){ $('v-' + x).hidden = (x !== v); });
  window.scrollTo(0, 0);
}
function tulisGalat(id, pesan){ const g = $(id); g.textContent = pesan || ''; g.hidden = !pesan; }
function adaAdmin(){ return !!ME && ME.role === 'admin'; }
function opsiList(k){ return k.opsi.split(',').map(function(s){ return s.trim(); }).filter(Boolean); }
function galatSesi(pesan){ const e = new Error(pesan || 'Akunmu diubah atau dinonaktifkan oleh admin. Silakan masuk lagi.'); e.kode = 'SESI'; return e; }
function infoMe(u){ return {username: u.u, role: u.role, ganti: !!u.ganti, hp: u.hash.slice(0, 12)}; }

/* =====================================================================
   GOOGLE: login (Google Identity Services) + Sheets API
   ===================================================================== */
function googleSiap(){ return !!(window.google && google.accounts && google.accounts.oauth2); }
function tungguGoogle(){
  return new Promise(function(ok, gagal){
    if (googleSiap()){ ok(); return; }
    let n = 0;
    const t = setInterval(function(){
      if (googleSiap()){ clearInterval(t); ok(); }
      else if (++n > 80){ clearInterval(t); gagal(new Error('Skrip Google belum termuat. Periksa koneksi internet, lalu muat ulang halaman.')); }
    }, 100);
  });
}
function pesanOAuth(kode){
  if (kode === 'access_denied') return 'Izin Google ditolak. Setujui akses ke Google Spreadsheet untuk melanjutkan.';
  if (kode === 'invalid_client' || kode === 'unauthorized_client') return 'Client ID tidak valid atau belum diizinkan untuk alamat situs ini.';
  return 'Gagal masuk ke Google' + (kode ? ' (' + kode + ')' : '') + '.';
}
function pesanPopup(e){
  const t = e && e.type;
  if (t === 'popup_closed') return 'Jendela Google ditutup sebelum selesai. Coba lagi.';
  if (t === 'popup_failed_to_open') return 'Browser memblokir jendela Google. Izinkan popup untuk situs ini, lalu coba lagi.';
  return 'Gagal membuka jendela Google.';
}
async function ambilEmail(){
  try{
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {headers: {Authorization: 'Bearer ' + tok.t}});
    if (r.ok){ const j = await r.json(); S.email = j.email || ''; }
  }catch(e){}
}
function simpanToken(){ simpanKe('sess', 'ed-token', {t: tok.t, exp: tok.exp, email: S.email}); }
function tokenValid(){ return !!tok.t && Date.now() < tok.exp - 60000; }
// Harus dipanggil langsung dari ketukan/klik pengguna (kalau tidak, popup bisa diblokir browser).
function mintaToken(prompt){
  if (tokenJanji) return tokenJanji;
  tokenJanji = tungguGoogle().then(function(){
    return new Promise(function(ok, gagal){
      if (!S.clientId){ gagal(new Error('Client ID belum diisi.')); return; }
      try{
        const klien = google.accounts.oauth2.initTokenClient({
          client_id: S.clientId,
          scope: SCOPE,
          callback: function(r){
            if (!r || r.error || !r.access_token){ gagal(new Error(pesanOAuth(r && r.error))); return; }
            tok.t = r.access_token;
            tok.exp = Date.now() + (Number(r.expires_in) || 3600) * 1000;
            ambilEmail().then(function(){ simpanToken(); ok(); });
          },
          error_callback: function(e){ gagal(new Error(pesanPopup(e))); }
        });
        klien.requestAccessToken({prompt: prompt});
      }catch(e){ gagal(e); }
    });
  });
  const lepas = function(){ tokenJanji = null; };
  tokenJanji.then(lepas, lepas);
  return tokenJanji;
}
function pastikanToken(){ return tokenValid() ? Promise.resolve() : mintaToken(''); }

function GalatApi(pesan, status, raw){ const e = new Error(pesan); e.status = status; e.raw = raw || ''; return e; }
function terjemah(status, msg){
  const m = String(msg || '');
  if (status === 403){
    if (/has not been used|is disabled|accessNotConfigured|SERVICE_DISABLED/i.test(m))
      return 'Google Sheets API belum diaktifkan di project Google Cloud. Buka Library, cari “Google Sheets API”, lalu klik Enable.';
    return 'Akun Google ' + (S.email ? '(' + S.email + ') ' : '') + 'tidak punya akses Editor ke spreadsheet ini. Bagikan spreadsheet sebagai Editor ke akun tersebut.';
  }
  if (status === 404) return 'Spreadsheet tidak ditemukan. Periksa linknya.';
  if (status === 429) return 'Terlalu banyak permintaan ke Google. Tunggu sebentar lalu coba lagi.';
  if (status === 401) return 'Sesi Google berakhir. Sambungkan ulang.';
  return m || ('Galat dari Google (' + status + ').');
}
async function api(metode, path, body, query){
  await pastikanToken();
  const url = 'https://sheets.googleapis.com/v4/' + path + (query ? '?' + query : '');
  for (let coba = 0; coba < 2; coba++){
    let res;
    const headers = {Authorization: 'Bearer ' + tok.t};
    if (body) headers['Content-Type'] = 'application/json';
    try{
      res = await fetch(url, {method: metode, headers: headers, body: body ? JSON.stringify(body) : undefined});
    }catch(e){ throw GalatApi('Tidak bisa terhubung ke Google. Periksa koneksi internet.', 0); }
    if (res.status === 401 && coba === 0){ tok.t = ''; await mintaToken(''); continue; }
    let data = null;
    try{ data = await res.json(); }catch(e){}
    if (!res.ok){
      const raw = (data && data.error && data.error.message) || '';
      throw GalatApi(terjemah(res.status, raw), res.status, raw);
    }
    return data || {};
  }
  throw GalatApi(terjemah(401), 401);
}
async function ambilMeta(id){
  const r = await api('GET', 'spreadsheets/' + id, null,
    'fields=' + encodeURIComponent('properties.title,sheets.properties(sheetId,title,hidden,gridProperties)'));
  return {judul: (r.properties || {}).title || 'Spreadsheet', sheets: (r.sheets || []).map(function(s){ return s.properties; })};
}
async function buatTab(id, nama, tersembunyi){
  const r = await api('POST', 'spreadsheets/' + id + ':batchUpdate',
    {requests: [{addSheet: {properties: {title: nama, hidden: !!tersembunyi}}}]});
  return r.replies && r.replies[0] && r.replies[0].addSheet && r.replies[0].addSheet.properties.sheetId;
}

/* =====================================================================
   PUSAT: akun + pengaturan, disimpan di tab "_pengaturan" (spreadsheet utama)
   Baris 1 = judul, baris 2 = "konfig" (JSON), baris 3+ = "user:<nama>" (JSON)
   ===================================================================== */
async function bootstrapPusat(){
  try{ await buatTab(S.homeId, TAB_PUSAT, true); }
  catch(e){ if (!(e.status === 400 && /already exists/i.test(e.raw))) throw e; }
  await api('PUT', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A1:C2'), {values: [
    ['kunci', 'nilai', 'Jangan diubah atau dihapus manual. Berisi akun dan pengaturan aplikasi Entri Data.'],
    ['konfig', JSON.stringify(konfigBawaan())]
  ]}, 'valueInputOption=RAW');
  await api('POST', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A1') + ':append',
    {values: [['user:' + ADMIN_USERNAME, jsonUser(adminBawaan())]]}, 'valueInputOption=RAW&insertDataOption=INSERT_ROWS');
}
async function bacaPusat(){
  const r = await api('GET', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A1:B600'), null, 'valueRenderOption=UNFORMATTED_VALUE');
  const rows = r.values || [];
  let konfig = null, konfigRow = 0;
  const users = [], seen = {};
  rows.forEach(function(row, i){
    const k = String(row[0] || ''), v = row[1];
    if (k === 'konfig'){
      try{ const j = bersihKonfig(JSON.parse(v)); if (j && !konfig){ konfig = j; konfigRow = i + 1; } }catch(e){}
    } else if (k.indexOf('user:') === 0 && v){
      try{
        const u = JSON.parse(v);
        if (u && u.u && u.hash && u.salt && !seen[u.u]){ seen[u.u] = 1; u._row = i + 1; users.push(u); }
      }catch(e){}
    }
  });
  S.users = users;
  S.konfigAda = !!konfig;
  S.konfigRow = konfigRow || 2;
  S.konfig = konfig || konfigBawaan();
}
async function tulisKonfig(k){
  await api('PUT', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A' + S.konfigRow + ':B' + S.konfigRow),
    {values: [['konfig', JSON.stringify(k)]]}, 'valueInputOption=RAW');
}
async function tulisUser(u){
  const sel = ['user:' + u.u, jsonUser(u)];
  if (u._row){
    await api('PUT', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A' + u._row + ':B' + u._row), {values: [sel]}, 'valueInputOption=RAW');
  } else {
    await api('POST', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A1') + ':append', {values: [sel]}, 'valueInputOption=RAW&insertDataOption=INSERT_ROWS');
  }
}
async function muatPusat(){
  const meta = await ambilMeta(S.homeId);
  S.judul = meta.judul;
  if (!meta.sheets.some(function(p){ return p.title === TAB_PUSAT; })) await bootstrapPusat();
  await bacaPusat();
  if (!S.konfigAda){ await tulisKonfig(S.konfig); S.konfigAda = true; }
  if (!S.users.some(function(u){ return u.role === 'admin' && u.aktif; })){
    // Tidak ada admin aktif (mis. barisnya terhapus): kembalikan akun admin bawaan.
    const lama = S.users.find(function(u){ return u.u === ADMIN_USERNAME; });
    const a = adminBawaan();
    if (lama) a._row = lama._row;
    await tulisUser(a);
    await bacaPusat();
  }
}
// Apakah akun yang sedang dipakai masih aktif dan passwordnya belum diganti dari perangkat lain?
async function sesiMasihSah(){
  const u = S.users.find(function(x){ return x.u === ME.username; });
  if (!u) return false;
  const r = await api('GET', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'B' + u._row), null, 'valueRenderOption=UNFORMATTED_VALUE');
  try{
    const d = JSON.parse(((r.values || [])[0] || [])[0]);
    return !!d && !!d.aktif && String(d.hash).slice(0, 12) === ME.hp;
  }catch(e){ return false; }
}

/* =====================================================================
   MENYIMPAN DATA KE SHEET TUJUAN
   ===================================================================== */
async function bacaHeader(id, nama){
  try{
    const r = await api('GET', 'spreadsheets/' + id + '/values/' + rng(nama, '1:1'), null, 'valueRenderOption=FORMATTED_VALUE');
    return {judul: ((r.values || [])[0] || []).map(function(x){ return String(x).trim(); }), baru: false};
  }catch(e){
    if (e.status === 400 && /Unable to parse range/i.test(e.raw)){ await buatTab(id, nama, false); return {judul: [], baru: true}; }
    throw e;
  }
}
// Kolom yang berformat "Teks biasa" (diambil sekali per sesi, lalu disimpan sementara)
async function kolomTeks(id, nama, judul){
  const kunci = id + '|' + nama + '|' + judul.join('\u0001');
  if (cacheFmt[kunci]) return cacheFmt[kunci];
  const set = {};
  try{
    const r = await api('GET', 'spreadsheets/' + id, null,
      'ranges=' + rng(nama, 'A2:' + colLetter(judul.length) + '2') +
      '&fields=' + encodeURIComponent('sheets.data.rowData.values.userEnteredFormat.numberFormat.type'));
    const d0 = ((r.sheets || [])[0] || {}).data;
    const row0 = ((d0 || [])[0] || {}).rowData;
    const vals = (((row0 || [])[0]) || {}).values || [];
    vals.forEach(function(c, i){
      const f = c && c.userEnteredFormat && c.userEnteredFormat.numberFormat;
      if (f && f.type === 'TEXT') set[i] = true;
    });
  }catch(e){}
  cacheFmt[kunci] = set;
  return set;
}
async function simpanBaris(items){
  await pastikanToken();                       // dipanggil pertama, selagi masih dalam ketukan pengguna
  const k = S.konfig;
  const id = ambilId(k.target) || S.homeId;
  const nama = k.sheet || 'Data';
  const semua = [];
  if (k.waktu) semua.push({nama: KOL_WAKTU, nilai: formatWaktu(new Date())});
  if (k.catatUser) semua.push({nama: KOL_USER, nilai: amankanTeks(ME.username)});
  items.forEach(function(it){ semua.push(it); });

  const hasil = await Promise.all([sesiMasihSah(), bacaHeader(id, nama)]);
  if (!hasil[0]) throw galatSesi();
  const hdr = hasil[1];

  const gab = gabungHeader(hdr.judul, semua);
  const judul = gab.judul;
  let info = null;
  async function infoSheet(){
    if (!info){ const m = await ambilMeta(id); info = m.sheets.find(function(p){ return p.title === nama; }) || {}; }
    return info;
  }
  if (gab.baru.length){
    const inf = await infoSheet();
    const kolomAda = inf.gridProperties && inf.gridProperties.columnCount;
    if (kolomAda && kolomAda < judul.length){
      await api('POST', 'spreadsheets/' + id + ':batchUpdate',
        {requests: [{appendDimension: {sheetId: inf.sheetId, dimension: 'COLUMNS', length: judul.length - kolomAda}}]});
    }
    await api('PUT', 'spreadsheets/' + id + '/values/' + rng(nama, colLetter(gab.awal + 1) + '1:' + colLetter(judul.length) + '1'),
      {values: [gab.baru]}, 'valueInputOption=RAW');
    if (gab.awal === 0 && inf.sheetId !== undefined){      // baris judul pertama: tebalkan & bekukan (tidak wajib)
      try{
        await api('POST', 'spreadsheets/' + id + ':batchUpdate', {requests: [
          {repeatCell: {range: {sheetId: inf.sheetId, startRowIndex: 0, endRowIndex: 1},
                        cell: {userEnteredFormat: {textFormat: {bold: true}}}, fields: 'userEnteredFormat.textFormat.bold'}},
          {updateSheetProperties: {properties: {sheetId: inf.sheetId, gridProperties: {frozenRowCount: 1}}, fields: 'gridProperties.frozenRowCount'}}
        ]});
      }catch(e){}
    }
  }

  const baris = susunBaris(judul, semua);
  lepasPetik(baris, await kolomTeks(id, nama, judul));
  const res = await api('POST', 'spreadsheets/' + id + '/values/' + rng(nama, 'A1') + ':append',
    {values: [baris]}, 'valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS');
  const m = String((res.updates || {}).updatedRange || '').match(/!\$?[A-Z]+(\d+)/);
  return {baris: m ? m[1] : '?', sheet: nama, kolomBaru: gab.baru.length, sheetBaru: hdr.baru};
}

/* =====================================================================
   LOGIN
   ===================================================================== */
function bacaGagal(){ return bacaDari('local', 'ed-gagal') || {}; }
function cekKunci(key){
  const all = bacaGagal(), g = all[key];
  if (g && g.n >= MAKS_GAGAL){
    const sisa = g.t + KUNCI_MS - Date.now();
    if (sisa > 0) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam ' + Math.ceil(sisa / 60000) + ' menit.');
    delete all[key]; simpanKe('local', 'ed-gagal', all);
  }
}
function catatGagal(key){
  const all = bacaGagal();
  const g = all[key] || {n: 0, t: 0};
  all[key] = {n: g.n + 1, t: Date.now()};
  simpanKe('local', 'ed-gagal', all);
}
function hapusGagal(key){ const all = bacaGagal(); delete all[key]; simpanKe('local', 'ed-gagal', all); }

function keLogin(pesan){
  ME = null; dirty = false;
  simpanKe('sess', 'ed-sesi', null);
  tampil('login');
  tulisGalat('galatLogin', pesan);
  $('lp').value = '';
  if ($('lu').value) $('lp').focus(); else $('lu').focus();
}
async function masukLogin(ev){
  ev.preventDefault();
  const un = $('lu').value.trim().toLowerCase(), pw = $('lp').value;
  if (!un || !pw){ tulisGalat('galatLogin', 'Isi username dan password.'); return; }
  const b = $('btnMasuk');
  tulisGalat('galatLogin', '');
  sibuk(b, true, 'Memeriksa…');
  try{
    cekKunci(un);
    if (!adaCrypto()) throw new Error('Browser ini tidak mendukung enkripsi yang dibutuhkan. Buka aplikasi lewat alamat https.');
    const coba = async function(){
      const u = S.users.find(function(x){ return x.u === un; });
      const h = await hashPw(pw, u ? u.salt : ADMIN_SALT);   // tetap dihitung walau username tidak ada
      return (u && samaPersis(h, u.hash)) ? u : null;
    };
    let u = await coba();
    if (!u && tokenValid()){            // mungkin akun baru/diubah di perangkat lain: muat ulang lalu cek sekali lagi
      try{ await bacaPusat(); u = await coba(); }catch(e){}
    }
    if (!u){ catatGagal(un); throw new Error('Username atau password salah.'); }
    if (!u.aktif) throw new Error('Akun ini dinonaktifkan. Hubungi admin.');
    hapusGagal(un);
    ME = infoMe(u);
    simpanKe('sess', 'ed-sesi', {u: u.u, h: ME.hp});
    $('lp').value = '';
    await masukApp();
  }catch(e){
    tulisGalat('galatLogin', e.message);
  }finally{
    sibuk(b, false);
  }
}
function keluar(){
  keLogin('');
  $('lu').value = '';
}

/* =====================================================================
   HUBUNGKAN GOOGLE (layar awal)
   ===================================================================== */
function simpanKoneksi(){ simpanKe('local', 'ed-koneksi', {clientId: S.clientId, url: S.homeUrl}); }
function renderSetup(pesan){
  tampil('setup');
  const ada = !!(S.clientId && S.homeId);
  $('cid').value = S.clientId;
  $('hurl').value = S.homeUrl;
  $('boxKoneksi').hidden = ada;
  $('btnUbahKoneksi').hidden = !ada;
  const r = $('ringkasKoneksi');
  r.hidden = !ada;
  r.textContent = ada ? 'Koneksi tersimpan di perangkat ini (spreadsheet …' + S.homeId.slice(-6) + '). Ketuk tombol di bawah, lalu pilih akun Google yang punya akses ke spreadsheet.' : '';
  let p = pesan || '';
  if (!p && location.protocol === 'file:') p = 'Halaman dibuka langsung dari file. Login Google hanya bisa dari alamat https (hosting) atau http://localhost.';
  tulisGalat('galatSetup', p);
}
async function hubungkanGoogle(){
  const cid = $('cid').value.trim(), url = $('hurl').value.trim();
  if (!/\.apps\.googleusercontent\.com$/.test(cid)){ tulisGalat('galatSetup', 'Client ID belum benar. Bentuknya berakhiran .apps.googleusercontent.com'); $('boxKoneksi').hidden = false; $('cid').focus(); return; }
  if (!ambilId(url)){ tulisGalat('galatSetup', 'Link spreadsheet belum benar. Salin link lengkap dari Google Sheets.'); $('boxKoneksi').hidden = false; $('hurl').focus(); return; }
  S.clientId = cid; S.homeUrl = url; S.homeId = ambilId(url);
  simpanKoneksi();
  const b = $('btnGoogle');
  tulisGalat('galatSetup', '');
  sibuk(b, true, 'Menghubungkan…');
  try{
    await mintaToken('');
    await masukPusat();
  }catch(e){
    renderSetup(e.message);
  }finally{
    sibuk(b, false);
  }
}
async function masukPusat(){
  await muatPusat();
  const sesi = bacaDari('sess', 'ed-sesi');
  const u = sesi && S.users.find(function(x){ return x.u === sesi.u && x.aktif && x.hash.slice(0, 12) === sesi.h; });
  if (u){ ME = infoMe(u); await masukApp(); }
  else keLogin('');
}

/* =====================================================================
   APLIKASI
   ===================================================================== */
function muatKeKerja(){
  const k = S.konfig;
  C = {urlInput: k.target || S.homeUrl, sheet: k.sheet || 'Data', waktu: !!k.waktu, catatUser: k.catatUser !== false};
  K = k.kolom.map(function(x){ return {id: uid(), nama: x.nama, tipe: x.tipe, opsi: x.opsi || ''}; });
  dirty = false;
  $('url').value = C.urlInput;
  $('sheet').value = C.sheet;
  $('waktu').checked = C.waktu;
  $('catatUser').checked = C.catatUser;
  $('infoSheet').hidden = true;
  $('linkSetup').hidden = true;
  const a = $('akunHint');
  if (S.email){ a.textContent = 'Terhubung dengan akun Google ' + S.email + '. Spreadsheet harus dibagikan sebagai Editor ke akun ini.'; a.hidden = false; }
  else a.hidden = true;
}
function tampilkanWajibGanti(){
  tampil('app');
  $('siapa').textContent = ME.username + ' · ' + (adaAdmin() ? 'Admin' : 'Pengguna');
  $('status').hidden = true;
  $('tabs').hidden = true;
  $('btnPass').hidden = true; $('btnSegar').hidden = true;
  $('btnKembali').hidden = true;
  $('wajibInfo').hidden = false;
  $('formPass').reset();
  panelTampil('akun');
}
async function masukApp(){
  if (ME.ganti){ tampilkanWajibGanti(); return; }
  tampil('app');
  $('siapa').textContent = ME.username + ' · ' + (adaAdmin() ? 'Admin' : 'Pengguna');
  $('status').hidden = false;
  $('btnPass').hidden = false; $('btnSegar').hidden = false;
  $('btnKembali').hidden = false;
  $('wajibInfo').hidden = true;
  $('tabs').hidden = !adaAdmin();
  muatKeKerja();
  S.judulTarget = '';
  renderStatus();
  renderKolom();
  await bukaTab('isi');
  if (S.konfig.target){
    try{ S.judulTarget = (await ambilMeta(ambilId(S.konfig.target))).judul; renderStatus(); }catch(e){}
  }
}
function renderStatus(){
  const p = $('status');
  const judul = S.konfig.target ? (S.judulTarget || 'spreadsheet tujuan') : (S.judul || 'Spreadsheet');
  const belum = adaAdmin() && dirty;
  p.className = 'pill ' + (belum ? 'warn' : 'ok');
  p.textContent = '✓ ' + judul + ' › ' + (S.konfig.sheet || 'Data') + (belum ? ' · belum disimpan' : '');
}
function markDirty(){ dirty = true; renderStatus(); }
function panelTampil(nama){
  ['isi', 'sheet', 'kolom', 'user', 'akun'].forEach(function(t){ $('p-' + t).hidden = (t !== nama); });
  document.querySelectorAll('[data-tab]').forEach(function(b){
    b.setAttribute('aria-selected', String(b.getAttribute('data-tab') === nama));
  });
  window.scrollTo(0, 0);
}
async function bukaTab(nama){
  if (!adaAdmin() && nama !== 'isi' && nama !== 'akun') nama = 'isi';
  if (nama === 'isi' && adaAdmin() && dirty){
    const err = validasiPengaturan();
    if (err){ toast(err, true); nama = 'kolom'; }
    else if (!(await simpanPengaturan())) return;
  }
  if (nama === 'isi') renderForm();
  if (nama === 'user') muatUser();
  if (nama === 'akun') $('formPass').reset();
  panelTampil(nama);
}
async function segarkan(){
  if (dirty){ toast('Simpan dulu pengaturan yang belum disimpan.', true); return; }
  const b = $('btnSegar');
  b.disabled = true;
  try{
    await muatPusat();
    const u = S.users.find(function(x){ return x.u === ME.username && x.aktif && x.hash.slice(0, 12) === ME.hp; });
    if (!u){ keLogin('Akunmu berubah atau dinonaktifkan. Silakan masuk lagi.'); return; }
    ME = infoMe(u);
    await masukApp();
    toast('✓ Pengaturan dimuat ulang');
  }catch(e){
    toast(e.message, true);
  }finally{
    b.disabled = false;
  }
}

/* ---------- pengaturan (admin) ---------- */
function validasiPengaturan(){
  const nm = C.sheet.trim() || 'Data';
  const es = cekNamaSheet(nm);
  if (es) return es;
  const u = C.urlInput.trim();
  if (u && !ambilId(u)) return 'Link spreadsheet belum benar. Salin link lengkap dari Google Sheets.';
  if (!K.length) return 'Tambahkan minimal satu kolom dulu.';
  const seen = {};
  for (let i = 0; i < K.length; i++){
    const k = K[i], n = k.nama.trim(), low = n.toLowerCase();
    if (!n) return 'Kolom ke-' + (i + 1) + ' belum diberi nama.';
    if (C.waktu && low === KOL_WAKTU.toLowerCase()) return '“Waktu Input” sudah dipakai pencatat waktu otomatis. Ganti nama kolom atau matikan pencatat waktu.';
    if (C.catatUser && low === KOL_USER.toLowerCase()) return '“Diinput Oleh” sudah dipakai pencatat pengguna otomatis. Ganti nama kolom atau matikan pencatatnya.';
    if (seen[low]) return 'Nama kolom “' + n + '” dobel.';
    seen[low] = true;
    if (k.tipe === 'pilihan' && !opsiList(k).length) return 'Kolom “' + n + '” (Pilihan) belum punya opsi.';
  }
  return '';
}
async function simpanPengaturan(){
  const err = validasiPengaturan();
  if (err){ toast(err, true); return false; }
  const idInput = ambilId(C.urlInput.trim());
  const target = (!idInput || idInput === S.homeId) ? '' : C.urlInput.trim();
  const payload = {
    v: 1, target: target, sheet: C.sheet.trim() || 'Data', waktu: C.waktu, catatUser: C.catatUser,
    kolom: K.map(function(k){ return {nama: k.nama.trim(), tipe: k.tipe, opsi: k.opsi.slice(0, 300)}; })
  };
  if (JSON.stringify(payload).length > 45000){ toast('Pengaturan terlalu besar. Kurangi kolom atau panjang opsi pilihan.', true); return false; }
  try{
    await pastikanToken();
    if (target !== S.konfig.target){        // tujuan berubah: pastikan bisa dibuka sebelum dipakai semua orang
      S.judulTarget = target ? (await ambilMeta(idInput)).judul : '';
    }
    await tulisKonfig(payload);
    S.konfig = payload;
    dirty = false;
    renderStatus();
    toast('✓ Pengaturan disimpan');
    return true;
  }catch(e){
    toast('Gagal menyimpan pengaturan: ' + e.message, true);
    return false;
  }
}
function tombol(label, aria, fn, mati){
  const b = el('button', {type: 'button', 'class': 'icon', 'aria-label': aria, title: aria, disabled: mati}, label);
  b.addEventListener('click', fn);
  return b;
}
function geser(i, d){
  const j = i + d;
  if (j < 0 || j >= K.length) return;
  const t = K[i]; K[i] = K[j]; K[j] = t;
  markDirty(); renderKolom();
}
function renderKolom(){
  const box = $('daftarKolom');
  box.replaceChildren();
  K.forEach(function(k, i){
    const nama = el('input', {type: 'text', value: k.nama, placeholder: 'Nama kolom, mis. Alamat',
                              'aria-label': 'Nama kolom ' + (i + 1), maxlength: '60', autocomplete: 'off'});
    nama.addEventListener('input', function(){ k.nama = nama.value; markDirty(); });

    const tipe = el('select', {'aria-label': 'Jenis isian kolom ' + (i + 1)});
    TIPE.forEach(function(t){ tipe.append(el('option', {value: t.v, selected: (t.v === k.tipe)}, t.l)); });
    tipe.addEventListener('change', function(){ k.tipe = tipe.value; markDirty(); renderKolom(); });

    const baris2 = el('div', {'class': 'baris2'},
      tipe,
      tombol('↑', 'Pindah ke atas',  function(){ geser(i, -1); }, i === 0),
      tombol('↓', 'Pindah ke bawah', function(){ geser(i, 1); },  i === K.length - 1),
      tombol('✕', 'Hapus kolom',     function(){ K.splice(i, 1); markDirty(); renderKolom(); })
    );
    const row = el('div', {'class': 'krow'}, nama, baris2);
    if (k.tipe === 'pilihan'){
      const o = el('input', {type: 'text', value: k.opsi, maxlength: '300', placeholder: 'Opsi dipisah koma, mis. Lunas, Belum', 'aria-label': 'Opsi pilihan'});
      o.addEventListener('input', function(){ k.opsi = o.value; markDirty(); });
      row.append(o);
    }
    box.append(row);
  });
  if (!K.length) box.append(el('p', {'class': 'hint'}, 'Belum ada kolom. Tekan “Tambah kolom”.'));
}

/* ---------- hubungkan & cek sheet tujuan (admin) ---------- */
function tebakTipe(nama){
  const n = nama.toLowerCase();
  if (/(^|\W)(hp|wa|telp|telepon|phone|nik|npwp|rekening|rek|kode|nomor|no)(\W|$)/.test(n)) return 'kode';
  if (/tanggal|tgl|date/.test(n)) return 'tanggal';
  if (/keterangan|catatan|alamat|note|deskripsi/.test(n)) return 'catatan';
  if (/jumlah|qty|harga|total|nilai|umur|usia|berat|biaya/.test(n)) return 'angka';
  return 'teks';
}
function pakaiHeader(headers){
  const lama = {};
  K.forEach(function(k){ lama[k.nama.trim().toLowerCase()] = k; });
  const seen = {};
  C.waktu = headers.some(function(h){ return h.toLowerCase() === KOL_WAKTU.toLowerCase(); });
  C.catatUser = headers.some(function(h){ return h.toLowerCase() === KOL_USER.toLowerCase(); });
  K = headers.filter(function(h){
    const low = h.toLowerCase();
    if (low === KOL_WAKTU.toLowerCase() || low === KOL_USER.toLowerCase() || seen[low]) return false;
    seen[low] = true; return true;
  }).map(function(h){
    const m = lama[h.toLowerCase()];
    return {id: uid(), nama: h, tipe: m ? m.tipe : tebakTipe(h), opsi: m ? m.opsi : ''};
  });
  $('waktu').checked = C.waktu;
  $('catatUser').checked = C.catatUser;
  markDirty(); renderKolom();
}
function tampilInfo(r, galat){
  const box = $('infoSheet');
  box.replaceChildren();
  box.hidden = false;
  if (galat){
    box.className = 'note warn';
    box.append(el('strong', {}, 'Gagal terhubung. '), galat);
    return;
  }
  box.className = 'note';
  const nama = C.sheet.trim() || 'Data';
  box.append(el('strong', {}, '✓ Terhubung ke “' + r.judul + '”. '));
  if (r.ada && r.headers.length){
    box.append('Sheet “' + nama + '” sudah punya ' + r.headers.length + ' kolom: ' + r.headers.join(', ') + '. Pakai kolom ini untuk form?');
    const pakai = el('button', {type: 'button', 'class': 'btn primary'}, 'Pakai kolom ini');
    const tetap = el('button', {type: 'button', 'class': 'btn ghost'}, 'Pakai kolom saya');
    const row = el('div', {'class': 'row'}, pakai, tetap);
    pakai.addEventListener('click', function(){
      pakaiHeader(r.headers); row.remove();
      box.append(' Kolom form disamakan dengan sheet. Jangan lupa Simpan pengaturan.');
    });
    tetap.addEventListener('click', function(){
      row.remove();
      box.append(' Kolom milikmu yang belum ada di sheet akan ditambahkan di sebelah kanan.');
    });
    box.append(row);
  } else if (r.ada){
    box.append('Sheet “' + nama + '” masih kosong. Judul kolom ditulis saat data pertama disimpan.');
  } else {
    box.append('Sheet “' + nama + '” belum ada. Akan dibuat otomatis saat data pertama disimpan.');
  }
}
async function hubungkanTarget(){
  const id = ambilId(C.urlInput.trim()) || S.homeId;
  const nama = C.sheet.trim() || 'Data';
  const es = cekNamaSheet(nama);
  if (es){ toast(es, true); return; }
  if (C.urlInput.trim() && !ambilId(C.urlInput.trim())){ toast('Link spreadsheet belum benar. Salin link lengkap dari Google Sheets.', true); $('url').focus(); return; }
  const b = $('btnHubung');
  sibuk(b, true, 'Memeriksa…');
  try{
    const meta = await ambilMeta(id);
    const ada = meta.sheets.some(function(p){ return p.title === nama; });
    let headers = [];
    if (ada){
      const r = await api('GET', 'spreadsheets/' + id + '/values/' + rng(nama, '1:1'), null, 'valueRenderOption=FORMATTED_VALUE');
      headers = ((r.values || [])[0] || []).map(function(x){ return String(x).trim(); }).filter(Boolean);
    }
    S.judulTarget = id === S.homeId ? '' : meta.judul;
    const dl = $('daftarSheet');
    dl.replaceChildren();
    meta.sheets.filter(function(p){ return !p.hidden && p.title !== TAB_PUSAT; }).forEach(function(p){ dl.append(el('option', {value: p.title})); });
    tampilInfo({judul: meta.judul, ada: ada, headers: headers});
  }catch(e){
    tampilInfo(null, e.message);
  }finally{
    sibuk(b, false);
  }
}
async function bagikanSetup(){
  const link = location.origin + location.pathname + '#setup=' + buatTokenSetup(S.clientId, S.homeId);
  const inp = $('linkSetup');
  inp.value = link; inp.hidden = false;
  inp.focus(); inp.select();
  try{ await navigator.clipboard.writeText(link); toast('✓ Tautan disalin'); }
  catch(e){ toast('Salin tautan dari kolom di bawah tombol.'); }
}

/* ---------- form isi data ---------- */
function renderForm(){
  const box = $('isian');
  const lama = {};
  K.forEach(function(k){ const e = $('f' + k.id); if (e) lama[k.id] = e.value; });
  box.replaceChildren();
  if (!K.length) box.append(el('p', {'class': 'hint'}, 'Belum ada kolom isian. Admin perlu menambahkannya di tab Kolom.'));

  K.forEach(function(k, i){
    const id = 'f' + k.id;
    const hint = i === K.length - 1 ? 'done' : 'next';
    let inp;
    if (k.tipe === 'catatan') inp = el('textarea', {id: id, rows: '3'});
    else if (k.tipe === 'tanggal') inp = el('input', {id: id, type: 'date'});
    else if (k.tipe === 'pilihan'){
      inp = el('select', {id: id});
      inp.append(el('option', {value: ''}, '— pilih —'));
      opsiList(k).forEach(function(o){ inp.append(el('option', {value: o}, o)); });
    }
    else if (k.tipe === 'angka') inp = el('input', {id: id, type: 'text', inputmode: 'decimal', autocomplete: 'off', enterkeyhint: hint});
    else if (k.tipe === 'kode')  inp = el('input', {id: id, type: 'text', inputmode: 'tel', autocomplete: 'off', enterkeyhint: hint});
    else inp = el('input', {id: id, type: 'text', autocomplete: 'off', enterkeyhint: hint});

    if (lama[k.id] !== undefined) inp.value = lama[k.id];

    const nol = el('small', {'class': 'zero', hidden: true}, "Angka 0 di depan dijaga: disimpan sebagai teks (otomatis diberi ' di depannya).");
    if (k.tipe === 'teks' || k.tipe === 'angka' || k.tipe === 'kode'){
      const cek = function(){ nol.hidden = !/^\s*0\d/.test(inp.value); };
      inp.addEventListener('input', cek); cek();
    }
    box.append(el('div', {'class': 'field'}, el('label', {'for': id}, k.nama), inp, nol));
  });
}
function kumpulkan(){
  const items = [];
  let ada = false;
  K.forEach(function(k){
    const raw = ($('f' + k.id).value || '').trim();
    if (raw) ada = true;
    const nilai = siapkan(k.tipe, raw);
    if (nilai === null){ const e = new Error('“' + k.nama + '” harus berupa angka (desimal boleh pakai koma atau titik).'); e.fid = 'f' + k.id; throw e; }
    items.push({nama: k.nama.trim(), nilai: nilai});
  });
  if (!ada){ const e = new Error('Isi minimal satu kolom dulu.'); e.fid = K.length ? 'f' + K[0].id : ''; throw e; }
  return items;
}
async function kirimData(ev){
  ev.preventDefault();
  if (sedangKirim || !K.length) return;
  let items;
  try{ items = kumpulkan(); }
  catch(e){ toast(e.message, true); if (e.fid && $(e.fid)) $(e.fid).focus(); return; }

  sedangKirim = true;
  const btn = $('btnSimpan');
  sibuk(btn, true, 'Menyimpan…');
  try{
    const r = await simpanBaris(items);
    jumlahSesi++;
    let pesan = '✓ Tersimpan di baris ' + r.baris;
    if (r.sheetBaru) pesan += ' · sheet “' + r.sheet + '” dibuat';
    else if (r.kolomBaru) pesan += ' · ' + r.kolomBaru + ' kolom baru ditambahkan';
    toast(pesan);
    const jam = new Date().toLocaleTimeString('id-ID', {hour: '2-digit', minute: '2-digit'});
    $('terakhir').textContent = 'Terakhir: baris ' + r.baris + ' di sheet “' + r.sheet + '” · ' + jam + ' · ' + jumlahSesi + ' entri sesi ini';
    $('form').reset();
    document.querySelectorAll('.zero').forEach(function(z){ z.hidden = true; });
    const pertama = $('f' + K[0].id);
    if (pertama) pertama.focus();
  }catch(e){
    if (e.kode === 'SESI') keLogin(e.message);
    else toast('Gagal menyimpan: ' + e.message, true);
  }finally{
    sedangKirim = false;
    sibuk(btn, false);
  }
}

/* ---------- kelola pengguna (admin) ---------- */
function infoPass(judul, pw){
  const b = $('infoUser');
  b.replaceChildren();
  b.className = 'note';
  b.hidden = false;
  b.append(el('strong', {}, judul + ' '), 'Password sementara: ', el('code', {'class': 'pw'}, pw),
    '. Berikan ke pengguna; mereka wajib menggantinya saat masuk pertama. Catat sekarang, password ini tidak ditampilkan lagi.');
}
async function ubahUser(username, aksi){
  try{
    await pastikanToken();
    await bacaPusat();
    const u = S.users.find(function(x){ return x.u === username; });
    if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (username === ME.username) throw new Error('Tidak bisa mengubah akun sendiri dari sini.');
    let pw = '';
    if (aksi === 'reset'){
      pw = passwordAcak(); u.salt = saltBaru(); u.hash = await hashPw(pw, u.salt); u.ganti = true;
      hapusGagal(username);
    } else if (aksi === 'nonaktif') u.aktif = false;
    else if (aksi === 'aktif'){ u.aktif = true; hapusGagal(username); }
    else if (aksi === 'jadiAdmin') u.role = 'admin';
    else if (aksi === 'jadiUser') u.role = 'user';
    if (aksi === 'hapus'){
      await api('POST', 'spreadsheets/' + S.homeId + '/values/' + rng(TAB_PUSAT, 'A' + u._row + ':B' + u._row) + ':clear', {});
    } else {
      await tulisUser(u);
    }
    if (pw) infoPass('Password “' + username + '” direset.', pw); else toast('✓ Perubahan disimpan');
    await muatUser();
  }catch(e){
    toast(e.message, true);
  }
}
function renderUser(list){
  const box = $('daftarUser');
  box.replaceChildren();
  list.slice().sort(function(a, b){ return a.u < b.u ? -1 : 1; }).forEach(function(u){
    const anda = u.u === ME.username;
    const kepala = el('div', {'class': 'uhead'}, el('strong', {}, u.u));
    if (anda) kepala.append(el('span', {'class': 'badge'}, 'Anda'));
    kepala.append(el('span', {'class': 'badge' + (u.role === 'admin' ? ' admin' : '')}, u.role === 'admin' ? 'Admin' : 'Pengguna'));
    if (!u.aktif) kepala.append(el('span', {'class': 'badge off'}, 'Nonaktif'));
    if (u.ganti) kepala.append(el('span', {'class': 'badge off'}, 'Belum ganti password'));
    const row = el('div', {'class': 'urow'}, kepala);
    if (!anda){
      const aksi = el('div', {'class': 'uaksi'});
      const bReset = el('button', {type: 'button', 'class': 'kecil'}, 'Reset password');
      duaKetuk(bReset, 'Reset password', function(){ ubahUser(u.u, 'reset'); });
      const bAktif = el('button', {type: 'button', 'class': 'kecil'}, u.aktif ? 'Nonaktifkan' : 'Aktifkan');
      bAktif.addEventListener('click', function(){ ubahUser(u.u, u.aktif ? 'nonaktif' : 'aktif'); });
      const bPeran = el('button', {type: 'button', 'class': 'kecil'}, u.role === 'admin' ? 'Jadikan pengguna' : 'Jadikan admin');
      bPeran.addEventListener('click', function(){ ubahUser(u.u, u.role === 'admin' ? 'jadiUser' : 'jadiAdmin'); });
      const bHapus = el('button', {type: 'button', 'class': 'kecil'}, 'Hapus');
      duaKetuk(bHapus, 'Hapus', function(){ ubahUser(u.u, 'hapus'); });
      aksi.append(bReset, bAktif, bPeran, bHapus);
      row.append(aksi);
    }
    box.append(row);
  });
}
async function muatUser(){
  const box = $('daftarUser');
  box.textContent = 'Memuat…';
  try{
    await pastikanToken();
    await bacaPusat();
    renderUser(S.users);
  }catch(e){
    box.textContent = 'Gagal memuat: ' + e.message;
  }
}
async function tambahUser(ev){
  ev.preventDefault();
  const un = $('nu').value.trim().toLowerCase(), pw0 = $('np').value, peran = $('nr').value === 'admin' ? 'admin' : 'user';
  if (!cekUsername(un)){ toast('Username 3–30 karakter: huruf kecil, angka, titik, strip, atau garis bawah.', true); $('nu').focus(); return; }
  if (pw0 && pw0.length < 8){ toast('Password minimal 8 karakter (atau kosongkan agar dibuatkan otomatis).', true); $('np').focus(); return; }
  if (pw0 && pw0.toLowerCase() === un){ toast('Password tidak boleh sama dengan username.', true); $('np').focus(); return; }
  const b = $('btnTambahUser');
  sibuk(b, true, 'Menambahkan…');
  try{
    await pastikanToken();
    await bacaPusat();
    if (S.users.some(function(x){ return x.u === un; })) throw new Error('Username “' + un + '” sudah dipakai.');
    const pw = pw0 || passwordAcak();
    const salt = saltBaru();
    await tulisUser({u: un, salt: salt, hash: await hashPw(pw, salt), role: peran, aktif: true, ganti: true, dibuat: new Date().toISOString()});
    infoPass('Pengguna “' + un + '” dibuat.', pw);
    $('formUser').reset();
    await muatUser();
  }catch(e){
    toast(e.message, true);
  }finally{
    sibuk(b, false);
  }
}

/* ---------- ganti password ---------- */
async function gantiPass(ev){
  ev.preventDefault();
  const lama = $('pl').value, baru = $('pb').value, ulang = $('pu').value;
  if (!lama || !baru){ toast('Isi password lama dan password baru.', true); return; }
  if (baru.length < 8){ toast('Password baru minimal 8 karakter.', true); $('pb').focus(); return; }
  if (baru.toLowerCase() === ME.username){ toast('Password tidak boleh sama dengan username.', true); $('pb').focus(); return; }
  if (baru !== ulang){ toast('Ulangi password baru belum sama.', true); $('pu').focus(); return; }
  if (baru === lama){ toast('Password baru harus berbeda dari yang lama.', true); $('pb').focus(); return; }
  const b = $('btnGanti');
  sibuk(b, true, 'Menyimpan…');
  try{
    await pastikanToken();
    await bacaPusat();
    const u = S.users.find(function(x){ return x.u === ME.username; });
    if (!u) throw galatSesi();
    if (!samaPersis(await hashPw(lama, u.salt), u.hash)) throw new Error('Password lama salah.');
    u.salt = saltBaru();
    u.hash = await hashPw(baru, u.salt);
    u.ganti = false;
    await tulisUser(u);
    ME = infoMe(u);
    simpanKe('sess', 'ed-sesi', {u: u.u, h: ME.hp});
    $('formPass').reset();
    toast('✓ Password diganti');
    await masukApp();
  }catch(e){
    if (e.kode === 'SESI') keLogin(e.message); else toast(e.message, true);
  }finally{
    sibuk(b, false);
  }
}

/* =====================================================================
   MULAI
   ===================================================================== */
function pasangEvent(){
  $('originInfo').textContent = location.origin;
  $('btnGoogle').addEventListener('click', hubungkanGoogle);
  $('btnUbahKoneksi').addEventListener('click', function(){ $('boxKoneksi').hidden = false; $('btnUbahKoneksi').hidden = true; $('ringkasKoneksi').hidden = true; $('cid').focus(); });
  $('btnKoneksi').addEventListener('click', function(){ renderSetup(''); });
  ['cid', 'hurl'].forEach(function(id){
    $(id).addEventListener('keydown', function(e){ if (e.key === 'Enter'){ e.preventDefault(); hubungkanGoogle(); } });
  });

  $('formLogin').addEventListener('submit', masukLogin);
  $('lihat').addEventListener('change', function(){ $('lp').type = $('lihat').checked ? 'text' : 'password'; });
  $('btnKeluar').addEventListener('click', keluar);
  $('btnSegar').addEventListener('click', segarkan);
  $('btnPass').addEventListener('click', function(){ bukaTab('akun'); });
  $('btnKembali').addEventListener('click', function(){ bukaTab('isi'); });
  $('formPass').addEventListener('submit', gantiPass);
  $('formUser').addEventListener('submit', tambahUser);

  $('url').addEventListener('input', function(){ C.urlInput = $('url').value; $('infoSheet').hidden = true; markDirty(); });
  $('sheet').addEventListener('input', function(){ C.sheet = $('sheet').value; $('infoSheet').hidden = true; markDirty(); });
  ['url', 'sheet'].forEach(function(id){
    $(id).addEventListener('keydown', function(e){ if (e.key === 'Enter'){ e.preventDefault(); hubungkanTarget(); } });
  });
  $('waktu').addEventListener('change', function(){ C.waktu = $('waktu').checked; markDirty(); });
  $('catatUser').addEventListener('change', function(){ C.catatUser = $('catatUser').checked; markDirty(); });
  $('btnHubung').addEventListener('click', hubungkanTarget);
  $('btnSimpanSheet').addEventListener('click', simpanPengaturan);
  $('btnLinkSetup').addEventListener('click', bagikanSetup);
  $('btnTambah').addEventListener('click', function(){
    K.push({id: uid(), nama: '', tipe: 'teks', opsi: ''});
    markDirty(); renderKolom();
    const semua = $('daftarKolom').querySelectorAll('.krow > input');
    if (semua.length) semua[semua.length - 1].focus();
  });
  $('btnSelesai').addEventListener('click', function(){ bukaTab('isi'); });
  document.querySelectorAll('[data-tab]').forEach(function(b){
    b.addEventListener('click', function(){ bukaTab(b.getAttribute('data-tab')); });
  });
  $('form').addEventListener('submit', kirimData);
  // Enter = pindah ke isian berikutnya; di isian terakhir = simpan
  $('isian').addEventListener('keydown', function(e){
    if (e.key !== 'Enter' || e.isComposing) return;
    const t = e.target.tagName;
    if (t === 'TEXTAREA' || t === 'SELECT') return;
    e.preventDefault();
    const semua = Array.prototype.slice.call($('isian').querySelectorAll('input,select,textarea'));
    const i = semua.indexOf(e.target);
    if (i > -1 && i < semua.length - 1) semua[i + 1].focus(); else $('btnSimpan').click();
  });
}

async function mulai(){
  pasangEvent();
  const k = bacaDari('local', 'ed-koneksi') || {};
  S.clientId = k.clientId || GOOGLE_CLIENT_ID || '';
  S.homeUrl = k.url || '';
  S.homeId = ambilId(S.homeUrl);

  // Tautan setup dari admin: #setup=...
  const m = location.hash.match(/^#setup=([A-Za-z0-9_-]+)/);
  const st = m ? bacaTokenSetup(m[1]) : null;
  if (st){
    S.clientId = st.clientId; S.homeId = st.sheetId; S.homeUrl = urlDariId(st.sheetId);
    simpanKoneksi();
  }
  if (m){ try{ history.replaceState(null, '', location.pathname + location.search); }catch(e){} }

  // Lanjutkan token Google yang masih berlaku (mis. setelah muat ulang halaman)
  const t = bacaDari('sess', 'ed-token');
  if (t && t.t && Date.now() < t.exp - 60000 && S.clientId && S.homeId){
    tok = {t: t.t, exp: t.exp};
    S.email = t.email || '';
    try{ await masukPusat(); return; }
    catch(e){ renderSetup(e.message); return; }
  }
  renderSetup('');
}
mulai();

})();
