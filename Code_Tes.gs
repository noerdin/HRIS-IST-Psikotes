// ================================================================
//  SISTEM TES IST ONLINE - PT JAPA INDOTAMA
//  File: Code_Tes.gs  (Project GAS TERPISAH dari HRIS)
//  Versi: 1.0
//
//  CARA PAKAI:
//  1. Buat Google Apps Script project BARU (bukan yang sama dengan HRIS)
//  2. Paste file ini sebagai Code.gs
//  3. Buat file HTML baru bernama "Tes" → paste Tes.html
//  4. Di baris SPREADSHEET_ID, isi dengan ID spreadsheet HRIS Anda
//  5. Deploy sebagai Web App → Execute as: Me → Access: Anyone
//  6. Dari HRIS, panggil buatTokenPeserta() untuk generate link tes
// ================================================================

// ── KONFIGURASI ──
// Sistem akan mencari spreadsheet secara otomatis.
// Jika masih gagal, isi SPREADSHEET_ID_MANUAL dengan ID spreadsheet HRIS Anda.
// Cara ambil ID: buka spreadsheet HRIS → lihat URL:
// https://docs.google.com/spreadsheets/d/[COPY_ID_INI]/edit
var SPREADSHEET_ID_MANUAL = '';  // ← Isi jika auto-detect gagal
var NAMA_SPREADSHEET      = 'PT Japa Indotama';

var _ssCache = null;

function getSpreadsheet() {
  if (_ssCache) return _ssCache;

  // Metode 1: Gunakan ID manual jika sudah diisi
  if (SPREADSHEET_ID_MANUAL && SPREADSHEET_ID_MANUAL.length > 10) {
    try {
      _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID_MANUAL);
      return _ssCache;
    } catch(e) {
      console.error('openById gagal: ' + e.message);
    }
  }

  // Metode 2: Cari berdasarkan nama di Drive
  try {
    var files = DriveApp.getFilesByName(NAMA_SPREADSHEET);
    while (files.hasNext()) {
      var file = files.next();
      // Pastikan file adalah Spreadsheet
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        _ssCache = SpreadsheetApp.openById(file.getId());
        return _ssCache;
      }
    }
  } catch(e) {
    console.error('DriveApp search gagal: ' + e.message);
  }

  // Metode 3: getActiveSpreadsheet (jika script terhubung ke spreadsheet)
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) { _ssCache = active; return _ssCache; }
  } catch(e) {}

  throw new Error(
    'Spreadsheet tidak ditemukan. Isi SPREADSHEET_ID_MANUAL di Code_Tes.gs baris 8.'
  );
}

// ── FUNGSI DEBUG — jalankan ini dari editor GAS untuk diagnosis ──
function debugKoneksi() {
  var hasil = { ok: false, pesan: '', sheetAda: false, tokenCount: 0, spreadsheetName: '' };
  try {
    var ss = getSpreadsheet();
    hasil.spreadsheetName = ss.getName();
    hasil.ok = true;

    var sheet = ss.getSheetByName(SHEET_TOKEN);
    if (sheet) {
      hasil.sheetAda    = true;
      hasil.tokenCount  = Math.max(0, sheet.getLastRow() - 1);
    } else {
      hasil.pesan = 'Sheet "' + SHEET_TOKEN + '" belum ada (akan dibuat saat validasi pertama).';
    }
  } catch(e) {
    hasil.pesan = e.message;
  }
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

var SHEET_TOKEN     = 'db_token_psikotes';
var SHEET_HASIL     = 'db_hasil_psikotes';

// Header sheet db_token_psikotes
var HEADERS_TOKEN = [
  'TOKEN',           // 0 — kode unik 16 karakter
  'KANDIDAT_ID',     // 1 — dari db_kandidat
  'MPP_ID',          // 2 — dari db_mpp
  'NAMA_PESERTA',    // 3
  'EMAIL_PESERTA',   // 4
  'POSISI',          // 5
  'TGL_DIBUAT',      // 6
  'TGL_EXPIRED',     // 7 — expired 7 hari
  'STATUS',          // 8 — Belum Mulai / Sedang Berlangsung / Selesai / Expired
  'TGL_MULAI',       // 9
  'TGL_SELESAI',     // 10
  'IP_ADDRESS',      // 11
  'TAB_SWITCH_COUNT',// 12 — berapa kali pindah tab
  'DIBUAT_OLEH'      // 13
];

// Header sheet db_hasil_psikotes
var HEADERS_HASIL = [
  'TOKEN',           // 0
  'KANDIDAT_ID',     // 1
  'MPP_ID',          // 2
  'NAMA_PESERTA',    // 3
  'USIA_PESERTA',    // 4
  'TGL_TES',         // 5
  // Raw Score per subtes
  'RW_SE',  'RW_WA',  'RW_AN',  'RW_GE',  'RW_RA',
  'RW_ZR',  'RW_FA',  'RW_WU',  'RW_ME',
  // Standard Score per subtes
  'SW_SE',  'SW_WA',  'SW_AN',  'SW_GE',  'SW_RA',
  'SW_ZR',  'SW_FA',  'SW_WU',  'SW_ME',
  // Total & IQ
  'TOTAL_RW',        // jumlah semua RW
  'TOTAL_SW',        // jumlah semua SW
  'IQ',              // nilai IQ
  'KATEGORI_IQ',     // Sangat Tinggi / Tinggi / Rata-rata / dll
  // Dominasi kecerdasan
  'DOMINASI',
  // Jawaban mentah (JSON string) untuk verifikasi
  'JAWABAN_JSON',
  // Log keamanan
  'TAB_SWITCH_COUNT',
  'WAKTU_PENGERJAAN_MENIT',
  'TIMESTAMP'
];

// ================================================================
// doGet() — Entry point Web App
// ================================================================
function doGet(e) {
  var token = e.parameter.token || '';

  // Halaman admin (generate token) — hanya untuk HRD
  if (e.parameter.admin === '1') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Admin IST - PT Japa Indotama')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('Tes')
    .setTitle('Tes IST - PT Japa Indotama')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ================================================================
// validasiToken() — Validasi token peserta sebelum tes dimulai
// Dipanggil oleh Tes.html saat pertama load
// ================================================================
function validasiToken(token) {
  // Selalu return object — tidak boleh throw, karena withFailureHandler
  // di sisi client tidak selalu terpicu untuk semua jenis error GAS
  try {
    if (!token || token.length < 8) {
      return { valid: false, pesan: 'Token tidak valid. Periksa link Anda.' };
    }

    var ss;
    try {
      ss = getSpreadsheet();
    } catch(ssErr) {
      return {
        valid: false,
        pesan: 'Konfigurasi sistem bermasalah: ' + ssErr.message +
               ' — Hubungi HRD PT Japa Indotama.'
      };
    }

    var sheet = ss.getSheetByName(SHEET_TOKEN);
    if (!sheet) {
      // Sheet belum ada = belum ada token yang dibuat
      return { valid: false, pesan: 'Token tidak ditemukan. Hubungi HRD untuk mendapatkan link tes yang valid.' };
    }

    if (sheet.getLastRow() <= 1) {
      return { valid: false, pesan: 'Belum ada token yang terdaftar. Hubungi HRD.' };
    }

    var data = sheet.getDataRange().getValues();
    var rowIdx   = -1;
    var tokenRow = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === token.toString().trim()) {
        rowIdx   = i + 1;
        tokenRow = data[i];
        break;
      }
    }

    if (!tokenRow) {
      return { valid: false, pesan: 'Token tidak ditemukan. Periksa link atau hubungi HRD.' };
    }

    var status = tokenRow[8].toString().trim();

    if (status === 'Selesai') {
      return { valid: false, pesan: 'Tes sudah selesai dikerjakan. Token hanya bisa digunakan satu kali.' };
    }
    if (status === 'Expired') {
      return { valid: false, pesan: 'Token sudah kadaluarsa. Hubungi HRD untuk token baru.' };
    }

    // Cek tanggal expired (kolom index 7)
    var expiredRaw = tokenRow[7];
    if (expiredRaw) {
      var expDate = (expiredRaw instanceof Date) ? expiredRaw : new Date(expiredRaw);
      if (!isNaN(expDate.getTime()) && new Date() > expDate) {
        sheet.getRange(rowIdx, 9).setValue('Expired');
        return { valid: false, pesan: 'Token sudah kadaluarsa. Hubungi HRD untuk token baru.' };
      }
    }

    // Update status → Sedang Berlangsung (hanya jika masih Belum Mulai)
    if (status === 'Belum Mulai') {
      try {
        sheet.getRange(rowIdx, 9).setValue('Sedang Berlangsung');
        sheet.getRange(rowIdx, 10).setValue(formatTimestampTes(new Date()));
      } catch(updateErr) {
        console.error('Gagal update status: ' + updateErr.message);
        // Tetap lanjutkan — jangan gagalkan validasi hanya karena update status
      }
    }

    return {
      valid:      true,
      nama:       tokenRow[3].toString(),
      kandidatId: tokenRow[1].toString(),
      mppId:      tokenRow[2].toString(),
      posisi:     tokenRow[5].toString(),
      rowIdx:     rowIdx,
      status:     status
    };

  } catch (err) {
    // Catch-all — pastikan selalu return, tidak pernah throw ke client
    console.error('validasiToken ERROR: ' + err.message);
    return {
      valid: false,
      pesan: 'Terjadi kesalahan sistem (' + err.message + '). Coba refresh halaman atau hubungi HRD.'
    };
  }
}

// ================================================================
// simpanHasil() — Simpan jawaban & hitung skor IST
// Dipanggil saat peserta submit tes terakhir (subtes 9)
// ================================================================
function simpanHasil(payload) {
  try {
    var ss          = getSpreadsheet();
    var sheetToken  = ss.getSheetByName(SHEET_TOKEN);
    var sheetHasil  = getOrCreateSheetTes(ss, SHEET_HASIL, HEADERS_HASIL, '#1a3a5c');

    var token = payload.token;

    // Cek token masih valid
    var tokenData = sheetToken.getDataRange().getValues();
    var tokenRowIdx = -1, tokenRow = null;
    for (var i = 1; i < tokenData.length; i++) {
      if (tokenData[i][0].toString() === token) {
        tokenRowIdx = i + 1;
        tokenRow    = tokenData[i];
        break;
      }
    }
    if (!tokenRow) return { status: 'error', pesan: 'Token tidak valid.' };
    if (tokenRow[8].toString() === 'Selesai') return { status: 'error', pesan: 'Tes sudah pernah disubmit.' };

    // ── Hitung Raw Score (RW) per subtes ──
    var jawaban  = payload.jawaban;   // object { 1: 'e', 2: 'c', ..., 176: 'd' }
    var usia     = parseInt(payload.usia) || 25;

    var rwSE = hitungRW_SE(jawaban);
    var rwWA = hitungRW_WA(jawaban);
    var rwAN = hitungRW_AN(jawaban);
    var rwGE = hitungRW_GE(jawaban);
    var rwRA = hitungRW_RA(jawaban);
    var rwZR = hitungRW_ZR(jawaban);
    var rwFA = hitungRW_FA(jawaban);
    var rwWU = hitungRW_WU(jawaban);
    var rwME = hitungRW_ME(jawaban);

    // ── Konversi RW → SW berdasarkan usia ──
    var sw = konversiSW(usia, rwSE, rwWA, rwAN, rwGE, rwRA, rwZR, rwFA, rwWU, rwME);

    var totalRW = rwSE + rwWA + rwAN + rwGE + rwRA + rwZR + rwFA + rwWU + rwME;
    var totalSW = sw.SE + sw.WA + sw.AN + sw.GE + sw.RA + sw.ZR + sw.FA + sw.WU + sw.ME;

    // ── Hitung IQ dari Total SW ──
    var iq       = hitungIQ(totalSW);
    var kategori = kategoriIQ(iq);
    var dominasi = hitungDominasi(sw);

    // ── Simpan ke sheet db_hasil_psikotes ──
    var now = new Date();
    var row = [
      token,
      tokenRow[1],  // KANDIDAT_ID
      tokenRow[2],  // MPP_ID
      tokenRow[3],  // NAMA
      usia,
      formatTanggalTes(now),
      rwSE, rwWA, rwAN, rwGE, rwRA, rwZR, rwFA, rwWU, rwME,
      sw.SE, sw.WA, sw.AN, sw.GE, sw.RA, sw.ZR, sw.FA, sw.WU, sw.ME,
      totalRW, totalSW, iq, kategori, dominasi,
      JSON.stringify(jawaban),
      payload.tabSwitchCount || 0,
      payload.durasiMenit    || 0,
      formatTimestampTes(now)
    ];
    sheetHasil.appendRow(row);
    applyRowFormatTes(sheetHasil, sheetHasil.getLastRow(), HEADERS_HASIL.length);

    // ── Update status token → Selesai ──
    sheetToken.getRange(tokenRowIdx, 9).setValue('Selesai');
    sheetToken.getRange(tokenRowIdx, 11).setValue(formatTimestampTes(now));
    sheetToken.getRange(tokenRowIdx, 13).setValue(payload.tabSwitchCount || 0);

    // ── Update kolom hasil psikotes di db_kandidat (integrasi HRIS) ──
    updateHasilDiKandidat(ss, tokenRow[1].toString(), iq, kategori, totalSW, rwSE, rwWA, rwAN, rwGE, rwRA, rwZR, rwFA, rwWU, rwME);

    return { status: 'ok', pesan: 'Tes selesai. Terima kasih telah mengikuti seleksi.' };
  } catch (err) {
    return { status: 'error', pesan: 'Gagal menyimpan hasil: ' + err.message };
  }
}

// ================================================================
// updateHasilDiKandidat() — Update kolom hasil psikotes di db_kandidat HRIS
// ================================================================
function updateHasilDiKandidat(ss, kandidatId, iq, kategori, totalSW, rwSE, rwWA, rwAN, rwGE, rwRA, rwZR, rwFA, rwWU, rwME) {
  try {
    var sheet = ss.getSheetByName('db_kandidat');
    if (!sheet || sheet.getLastRow() <= 1) return;

    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString() === kandidatId.toString()) {
        var rowNum = i + 2;
        // Kolom 20 = TGL PSIKOTES, 21 = HASIL PSIKOTES, 22 = CATATAN PSIKOTES
        sheet.getRange(rowNum, 20).setValue(formatTanggalTes(new Date()));
        sheet.getRange(rowNum, 21).setValue('Lanjut');  // Default Lanjut, HRD bisa ubah
        sheet.getRange(rowNum, 22).setValue(
          'IQ: ' + iq + ' (' + kategori + ') | SW Total: ' + totalSW +
          ' | SE:'+rwSE+' WA:'+rwWA+' AN:'+rwAN+' GE:'+rwGE+' RA:'+rwRA+
          ' ZR:'+rwZR+' FA:'+rwFA+' WU:'+rwWU+' ME:'+rwME
        );
        // Update tahapan kandidat
        sheet.getRange(rowNum, 11).setValue('Interview HRD');
        sheet.getRange(rowNum, 34).setValue(formatTimestampTes(new Date()));
        break;
      }
    }
  } catch (e) {
    console.error('updateHasilDiKandidat ERROR: ' + e.message);
  }
}

// ================================================================
// autoSaveProgress() — Auto-save jawaban sementara (tiap 30 detik)
// Untuk recovery jika peserta refresh/close accidental
// ================================================================
function autoSaveProgress(payload) {
  try {
    // Simpan ke cache sementara (PropertiesService per token)
    var props = PropertiesService.getScriptProperties();
    props.setProperty('progress_' + payload.token, JSON.stringify({
      jawaban:  payload.jawaban,
      subtes:   payload.subtesAktif,
      saved:    new Date().getTime()
    }));
    return { status: 'ok' };
  } catch(e) {
    return { status: 'error' };
  }
}

// ================================================================
// getProgress() — Ambil progress tersimpan (jika ada)
// ================================================================
function getProgress(token) {
  try {
    var props = PropertiesService.getScriptProperties();
    var raw   = props.getProperty('progress_' + token);
    if (!raw) return null;
    var data  = JSON.parse(raw);
    // Progress expired setelah 2 jam
    if (new Date().getTime() - data.saved > 7200000) {
      props.deleteProperty('progress_' + token);
      return null;
    }
    return data;
  } catch(e) {
    return null;
  }
}

// ================================================================
// buatTokenPeserta() — Generate token baru untuk kandidat
// Dipanggil dari sistem HRIS (Code.gs utama) atau dari Admin panel
// ================================================================
function buatTokenPeserta(dataForm) {
  try {
    var ss    = getSpreadsheet();
    var sheet = getOrCreateSheetTes(ss, SHEET_TOKEN, HEADERS_TOKEN, '#132a47');

    var token   = generateToken();
    var now     = new Date();
    var expired = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // +7 hari

    var row = [
      token,
      dataForm.kandidatId  || '',
      dataForm.mppId       || '',
      dataForm.nama        || '',
      dataForm.email       || '',
      dataForm.posisi      || '',
      formatTimestampTes(now),
      formatTimestampTes(expired),
      'Belum Mulai',
      '', '', '', 0,
      'Admin HR'
    ];
    sheet.appendRow(row);
    applyRowFormatTes(sheet, sheet.getLastRow(), HEADERS_TOKEN.length);

    // Ambil URL Web App ini
    var url = ScriptApp.getService().getUrl() + '?token=' + token;

    return {
      status:  'success',
      token:   token,
      url:     url,
      expired: formatTanggalTes(expired),
      nama:    dataForm.nama
    };
  } catch (err) {
    return { status: 'error', pesan: err.message };
  }
}

// ================================================================
// getDaftarToken() — Ambil semua token (untuk admin panel HRIS)
// ================================================================
function getDaftarToken() {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_TOKEN);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow()-1, HEADERS_TOKEN.length).getValues();
    return data.map(function(r) {
      return r.map(function(c) {
        if (c instanceof Date) return formatTanggalTes(c);
        return c;
      });
    }).filter(function(r){ return r[0] !== ''; });
  } catch(e) { return []; }
}

// ================================================================
//  SCORING FUNCTIONS
// ================================================================

// ── Kunci Jawaban ──
var KUNCI = {
  // SE (1-20): pilihan ganda
  SE:  [null,'E','C','D','D','D','B','C','A','E','B','C','D','D','E','C','A','B','B','C','B'],
  // WA (21-40): pilihan ganda
  WA:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        'A','A','D','C','C','C','C','D','D','A','C','A','A','B','C','A','D','E','B','C'],
  // AN (41-60): pilihan ganda
  AN:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        'C','E','D','D','D','A','D','B','E','D','C','C','C','C','D','C','C','D','E','E'],
  // RA (77-96): jawaban angka
  RA:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        35,280,205,26,30,70,45,50,84,78,19,6,75,90,120,17,36,5,48,1],
  // ZR (97-116): jawaban angka
  ZR:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        27,25,27,15,46,10,42,7,5,14,8,14,45,63,12,80,14,12,63,10],
  // FA (117-136): pilihan ganda
  FA:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        'A','C','B','A','D','B','C','E','D','A','E','A','D','D','C','B','B','A','C','A'],
  // WU (137-156): pilihan ganda
  WU:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        'A','C','D','E','A','C','D','C','E','A','B','D','E','B','D','B','A','E','B','C'],
  // ME (157-176): pilihan ganda
  ME:  [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,
        'D','E','B','A','C','A','D','E','C','B','B','A','E','C','D','B','E','A','C','D']
};

// Kunci jawaban GE (61-76) — scoring bertingkat (skor 2, 1, atau 0)
var KUNCI_GE = {
  61: { s2: ['bunga','kembang','perdu'], s1: ['tumbuh-tumbuhan','tangkai','harum'] },
  62: { s2: ['alat indera','indera','panca indera'], s1: ['organ','alat tubuh'] },
  63: { s2: ['hablur','kristal','zat arang'], s1: ['berkilauan','mengkilat','bening'] },
  64: { s2: ['musim'], s1: ['cuaca','iklim'] },
  65: { s2: ['pembawa berita','alat perhubungan'], s1: ['telekomunikasi','perhubungan','komunikasi'] },
  66: { s2: ['alat optik','optik'], s1: ['lensa','alat melihat','melihat'] },
  67: { s2: ['alat pencernaan'], s1: ['jalan makanan','perut','isi perut','pencernaan makanan'] },
  68: { s2: ['jumlah','kuantitas','penyebut jumlah'], s1: ['mengukur','ukuran','penyertaan jumlah'] },
  69: { s2: ['bibit','bakal','embrio','alat pembiak'], s1: ['sel','pembiakan','keturunan','permulaan penghidupan'] },
  70: { s2: ['simbol','lambang'], s1: ['tanda','nama','tanda pengenal'] },
  71: { s2: ['makhluk hidup','organism'], s1: ['makhluk','tumbuh','biologi','hayat','ilmu hayat'] },
  72: { s2: ['wadah','tempat pengisi','tempat penyimpan'], s1: ['alat','tempat sesuatu','tempat'] },
  73: { s2: ['pengertian waktu','batas'], s1: ['waktu','lamanya','masa','saat','kata waktu'] },
  74: { s2: ['kata sifat','watak','sifat karakter'], s1: ['sifat','karakter'] },
  75: { s2: ['regulator harga','pengertian ekonomi'], s1: ['dagang','niaga','jual beli','pembelian','penjualan'] },
  76: { s2: ['pengertian ruang','penyebut ruang'], s1: ['arah','tempat','ruang','letak','penunjuk tempat'] }
};

function hitungRW_SE(j) {
  var skor = 0;
  for (var n = 1; n <= 20; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.SE[n]) skor++;
  }
  return skor;
}

function hitungRW_WA(j) {
  var skor = 0;
  for (var n = 21; n <= 40; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.WA[n]) skor++;
  }
  return skor;
}

function hitungRW_AN(j) {
  var skor = 0;
  for (var n = 41; n <= 60; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.AN[n]) skor++;
  }
  return skor;
}

function hitungRW_GE(j) {
  var skor = 0;
  for (var n = 61; n <= 76; n++) {
    var jawab = (j[n] || '').toString().toLowerCase().trim();
    if (!jawab) continue;
    var kunci = KUNCI_GE[n];
    if (!kunci) continue;
    // Cek skor 2 dulu
    var dapat2 = kunci.s2.some(function(k) { return jawab.includes(k.toLowerCase()); });
    if (dapat2) { skor += 2; continue; }
    var dapat1 = kunci.s1.some(function(k) { return jawab.includes(k.toLowerCase()); });
    if (dapat1) { skor += 1; }
  }
  return skor;
}

function hitungRW_RA(j) {
  var skor = 0;
  for (var n = 77; n <= 96; n++) {
    var jawab = parseFloat((j[n] || '').toString().replace(/[^0-9.]/g,''));
    if (!isNaN(jawab) && Math.abs(jawab - KUNCI.RA[n]) < 0.5) skor++;
  }
  return skor;
}

function hitungRW_ZR(j) {
  var skor = 0;
  for (var n = 97; n <= 116; n++) {
    var jawab = parseFloat((j[n] || '').toString().replace(/[^0-9.,-]/g,''));
    if (!isNaN(jawab) && Math.abs(jawab - KUNCI.ZR[n]) < 0.5) skor++;
  }
  return skor;
}

function hitungRW_FA(j) {
  var skor = 0;
  for (var n = 117; n <= 136; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.FA[n]) skor++;
  }
  return skor;
}

function hitungRW_WU(j) {
  var skor = 0;
  for (var n = 137; n <= 156; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.WU[n]) skor++;
  }
  return skor;
}

function hitungRW_ME(j) {
  var skor = 0;
  for (var n = 157; n <= 176; n++) {
    if (j[n] && j[n].toString().toUpperCase() === KUNCI.ME[n]) skor++;
  }
  return skor;
}

// ================================================================
// konversiSW() — Konversi RW → Standard Score berdasarkan usia
// Menggunakan tabel norma dari file APLIKASI_Skoring_IST.xlsx
// ================================================================
function konversiSW(usia, rwSE, rwWA, rwAN, rwGE, rwRA, rwZR, rwFA, rwWU, rwME) {
  // Pilih tabel norma berdasarkan kelompok usia
  var tabel = getTabelNorma(usia);

  return {
    SE: lookupSW(tabel.SE, rwSE),
    WA: lookupSW(tabel.WA, rwWA),
    AN: lookupSW(tabel.AN, rwAN),
    GE: lookupSW_GE(tabel.GE, rwGE),   // GE menggunakan skala berbeda (maks 32)
    RA: lookupSW(tabel.RA, rwRA),
    ZR: lookupSW(tabel.ZR, rwZR),
    FA: lookupSW(tabel.FA, rwFA),
    WU: lookupSW(tabel.WU, rwWU),
    ME: lookupSW(tabel.ME, rwME)
  };
}

// Lookup nilai SW dari tabel [rw, sw] menggunakan interpolasi
function lookupSW(tabel, rw) {
  rw = Math.max(0, Math.min(rw, tabel.length - 1));
  return tabel[rw] || tabel[0];
}

// GE menggunakan skala konversi khusus (RW bisa > 20 karena skor 0/1/2)
function lookupSW_GE(tabel, rwGE) {
  // Konversi skor GE (0-32) → SW menggunakan tabel konversi GE
  var konversiGE = [
    1, 1, 2, 3, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12,
    13, 13, 14, 14, 15, 15, 16, 17, 18, 19, 20, 20
  ];
  var idx = Math.max(0, Math.min(rwGE, 32));
  var konvGE = konversiGE[idx] || 1;
  return tabel[konvGE] || tabel[0];
}

// ── Tabel Norma per Kelompok Usia (dari file APLIKASI_Skoring_IST.xlsx) ──
// Format: array[rw] = sw, index 0 = RW 0, index 20 = RW 20
function getTabelNorma(usia) {
  if (usia >= 21 && usia <= 25) return NORMA_21_25;
  if (usia >= 26 && usia <= 30) return NORMA_26_30;
  if (usia >= 31 && usia <= 35) return NORMA_31_35;
  if (usia >= 36 && usia <= 40) return NORMA_36_40;
  if (usia >= 41 && usia <= 45) return NORMA_41_45;
  if (usia >= 46 && usia <= 50) return NORMA_46_50;
  // Default usia 21-25
  return NORMA_21_25;
}

// ── Tabel norma usia 21-25 ──
// Sumber: Sheet "USIA 21 - 25" dari APLIKASI_Skoring_IST.xlsx
// Index = RW (0..20), nilai = SW
var NORMA_21_25 = {
  SE:  [68,71,74,76,79,82,85,88,91,94,97,100,103,106,109,112,115,118,121,124,126],
  WA:  [63,66,70,74,77,81,84,88,91,95,99,102,106,109,113,116,120,124,127,131,134],
  AN:  [76,78,81,83,86,88,91,93,96,98,101,103,106,108,111,113,116,118,121,123,126],
  GE:  [69,72,75,78,81,83,86,89,92,94,97,100,103,106,108,111,114,117,119,122,125], // index = konvGE (0..20)
  RA:  [74,77,79,82,85,88,91,94,97,99,102,105,108,111,114,117,119,122,125,128,131],
  ZR:  [77,80,82,84,87,89,91,94,96,99,101,103,106,108,110,113,115,118,120,122,125],
  FA:  [70,73,76,79,81,84,87,90,93,96,99,101,104,107,110,113,116,119,121,124,127],
  WU:  [72,75,77,80,83,86,89,92,95,97,100,103,106,109,112,115,117,120,123,126,129],
  ME:  [75,77,80,82,84,87,89,91,94,96,98,101,102,105,108,110,112,115,117,119,122]
};

// ── Tabel norma usia 26-30 ──
var NORMA_26_30 = {
  SE:  [66,69,72,75,78,81,84,87,90,93,96,99,102,105,108,112,115,118,121,124,127],
  WA:  [66,69,73,76,79,83,86,89,93,96,99,103,106,109,113,116,119,123,126,129,133],
  AN:  [78,80,83,85,87,90,92,95,97,99,102,104,106,109,111,114,116,118,121,123,125],
  GE:  [69,71,74,77,80,83,85,88,91,94,96,99,102,105,108,110,113,116,119,121,124],
  RA:  [74,77,79,82,85,88,91,94,97,99,102,105,108,111,114,117,119,122,125,128,131],
  ZR:  [79,81,83,86,88,90,93,95,97,100,102,104,107,109,111,113,116,118,120,123,125],
  FA:  [71,73,76,79,82,85,88,91,93,96,99,102,105,108,111,113,116,119,122,125,128],
  WU:  [72,75,78,81,84,87,90,93,96,99,101,104,107,110,113,116,119,122,125,128,131],
  ME:  [77,80,82,84,86,89,91,93,95,98,100,102,105,107,109,111,114,116,118,120,123]
};

// ── Tabel norma usia 31-35 (gunakan 26-30 sebagai approx jika tidak ada data) ──
var NORMA_31_35 = NORMA_26_30;
var NORMA_36_40 = NORMA_26_30;
var NORMA_41_45 = NORMA_26_30;
var NORMA_46_50 = NORMA_26_30;

// ================================================================
// hitungIQ() — Konversi Total SW → IQ
// Berdasarkan Norma IQ dari sheet "Norma IQ"
// ================================================================
function hitungIQ(totalSW) {
  // Tabel: [batas_bawah_SW, IQ]
  // (dari tabel WS=IQ di sheet Norma IQ)
  if (totalSW >= 171) return 132;
  if (totalSW >= 161) return 128;
  if (totalSW >= 151) return 124;
  if (totalSW >= 141) return 121;
  if (totalSW >= 131) return 117;
  if (totalSW >= 121) return 113;
  if (totalSW >= 111) return 110;
  if (totalSW >= 101) return 106;
  if (totalSW >= 91)  return 102;
  if (totalSW >= 81)  return 98;
  if (totalSW >= 71)  return 94;
  if (totalSW >= 61)  return 90;
  if (totalSW >= 51)  return 87;
  if (totalSW >= 41)  return 83;
  if (totalSW >= 31)  return 79;
  if (totalSW >= 21)  return 75;
  if (totalSW >= 11)  return 71;
  return 67;
}

function kategoriIQ(iq) {
  if (iq >= 130) return 'Sangat Superior';
  if (iq >= 120) return 'Superior';
  if (iq >= 110) return 'Di Atas Rata-rata';
  if (iq >= 90)  return 'Rata-rata';
  if (iq >= 80)  return 'Di Bawah Rata-rata';
  if (iq >= 70)  return 'Borderline';
  return 'Rendah';
}

function hitungDominasi(sw) {
  // Kelompok: Verbal (SE,WA,AN), Numerik (RA,ZR), Figural (FA,WU), Memori (ME), Umum (GE)
  var verbal   = (sw.SE + sw.WA + sw.AN) / 3;
  var numerik  = (sw.RA + sw.ZR) / 2;
  var figural  = (sw.FA + sw.WU) / 2;
  var memori   = sw.ME;
  var umum     = sw.GE;

  var max = Math.max(verbal, numerik, figural, memori, umum);
  if (max === verbal)  return 'Verbal';
  if (max === numerik) return 'Numerik';
  if (max === figural) return 'Figural/Spasial';
  if (max === memori)  return 'Memori';
  return 'Umum';
}

// ================================================================
//  HELPERS
// ================================================================
function generateToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var token = '';
  for (var i = 0; i < 16; i++) {
    if (i > 0 && i % 4 === 0) token += '-';
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;  // Format: XXXX-XXXX-XXXX-XXXX
}

function getOrCreateSheetTes(ss, name, headers, bg) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground(bg||'#1a3a5c').setFontColor('#fff').setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(9);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function applyRowFormatTes(sheet, rowNum, colCount) {
  try {
    sheet.getRange(rowNum, 1, 1, colCount)
         .setBorder(true,true,true,true,true,true,'#e2e8f0',SpreadsheetApp.BorderStyle.SOLID);
  } catch(e) {}
}

function formatTanggalTes(date) {
  if (!date || !(date instanceof Date)) return '';
  var d = date.getDate().toString().padStart(2,'0');
  var m = (date.getMonth()+1).toString().padStart(2,'0');
  var y = date.getFullYear();
  return d+'/'+m+'/'+y;
}

function formatTimestampTes(date) {
  return Utilities.formatDate(date, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
}
