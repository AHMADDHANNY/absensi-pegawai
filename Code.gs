/**
 * APLIKASI PRESENSI PEGAWAI - BACKEND CORE (FIXED v4)
 *
 * CHANGELOG v4:
 * [FIX #1] allPermits selalu dikirim di getDashboardData() — bukan hanya saat ada attendance
 * [FIX #2] generateId() pakai max ID + 1, bukan lastRow — cegah duplikat saat baris dihapus
 * [FIX #3] Password di-hash dengan SHA-256 via Utilities.computeDigest()
 * [FIX #4] formatTgl() — normalisasi tanggal konsisten, cegah bug auto-convert Sheets
 * [FIX #5] Validasi input di addUser, editUser, submitPermit, submitAttendance
 *
 * Struktur kolom Attendance:
 * [0]id  [1]userId  [2]tgl  [3]jamIn  [4]jamOut
 * [5]lat  [6]lng  [7]latOut  [8]lngOut  [9]status  [10]photoIn  [11]photoOut
 *
 * Sheet Settings — key yang digunakan:
 *   office_lat   | Latitude kantor
 *   office_lng   | Longitude kantor
 *   radius       | Radius absen (meter)
 */

// ─────────────────────────────────────────────
//  INISIALISASI
// ─────────────────────────────────────────────

function getSS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan.');
  return ss;
}

function getSheetSafe(name) {
  var sheet = getSS().getSheetByName(name);
  if (!sheet) {
    throw new Error(
      'Sheet "' + name + '" tidak ditemukan. ' +
      'Jalankan setupSheets() terlebih dahulu dari Apps Script editor.'
    );
  }
  return sheet;
}

// ─────────────────────────────────────────────
//  [FIX #4] NORMALISASI TANGGAL
//  Google Sheets kadang auto-convert string "2025-01-15" → Date object.
//  Fungsi ini memastikan semua tanggal selalu jadi string "yyyy-MM-dd".
// ─────────────────────────────────────────────

function formatTgl(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'GMT+7', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // Jika sudah format yyyy-MM-dd, langsung return
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Coba parse ulang
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd');
  } catch (e) {}
  return s;
}

// ─────────────────────────────────────────────
//  [FIX #3] HASH PASSWORD (SHA-256)
//  Menggantikan password plaintext.
//  PENTING: Jalankan migratePasswordsToHash() SEKALI dari editor
//           untuk mengenkripsi password lama di sheet.
// ─────────────────────────────────────────────

function hashPassword(plain) {
  if (!plain) return '';
  var bytes   = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain, Utilities.Charset.UTF_8);
  var hexStr  = bytes.map(function(b) {
    var hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
  return hexStr;
}

/**
 * Jalankan SEKALI dari editor Apps Script untuk migrasi password lama → hash.
 * Setelah selesai, fungsi ini tidak perlu dijalankan lagi.
 */
function migratePasswordsToHash() {
  var col   = getUserColumnMap();
  var sheet = getSheetSafe('Users');
  var data  = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var raw = String(data[i][col.password]);
    // Skip yang sudah hash (64 karakter hex)
    if (/^[0-9a-f]{64}$/.test(raw)) continue;
    sheet.getRange(i + 1, col.password + 1).setValue(hashPassword(raw));
    count++;
  }
  SpreadsheetApp.getUi().alert('Migrasi selesai: ' + count + ' password dienkripsi.');
}

// ─────────────────────────────────────────────
//  [FIX #2] GENERATE ID AMAN
//  Pakai max(ID) + 1 bukan lastRow,
//  sehingga hapus baris tidak menyebabkan ID duplikat.
// ─────────────────────────────────────────────

function generateId(sheetName) {
  var sheet = getSS().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return 1;
  var data = sheet.getDataRange().getValues();
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var id = parseInt(data[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

// ─────────────────────────────────────────────
//  AUTO-SETUP
// ─────────────────────────────────────────────

function setupSheets() {
  var ss = getSS();

  var sheetDefs = [
    { name: 'Users',      headers: ['id_user','nama','jabatan','nomor_hp','password','role'] },
    { name: 'Attendance', headers: ['id_absen','id_user','tanggal','jam_masuk','jam_keluar','lat_in','lng_in','lat_out','lng_out','status','foto masuk','foto keluar'] },
    { name: 'Permits',    headers: ['id_izin','id_user','nama','tanggal','jenis','alasan','lampiran','status_app'] },
    { name: 'Settings',   headers: ['key','value'] }
  ];

  sheetDefs.forEach(function (def) {
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.appendRow(def.headers);
      sheet.getRange(1, 1, 1, def.headers.length).setFontWeight('bold');
    }
  });

  var settingsSheet = ss.getSheetByName('Settings');
  var existingKeys  = settingsSheet.getDataRange().getValues()
                        .map(function (r) { return String(r[0]).toLowerCase(); });
  var defaults = [
    ['office_lat', 0],
    ['office_lng', 0],
    ['radius',     100],
  ];
  defaults.forEach(function (d) {
    if (existingKeys.indexOf(d[0]) === -1) settingsSheet.appendRow(d);
  });

  // Password admin default sudah langsung di-hash
  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([1, 'Administrator', 'Admin', '08000000000', hashPassword('admin123'), 'admin']);
  }

  return { success: true, message: 'Setup selesai!' };
}

// ─────────────────────────────────────────────
//  ENTRY POINT
// ─────────────────────────────────────────────

function doGet() {
  try { setupSheets(); } catch (e) {}
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Presensi Solider')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─────────────────────────────────────────────
//  FOLDER GOOGLE DRIVE
// ─────────────────────────────────────────────

var DRIVE_FOLDER_NAME = "FOTO ABSENSI";

// ─────────────────────────────────────────────
//  PEMETAAN KOLOM USER
// ─────────────────────────────────────────────

function getUserColumnMap() {
  const sheet     = getSheetSafe('Users');
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalize = function (s) { return String(s).trim().toLowerCase(); };

  const aliases = {
    id:       ['id','id_user','iduser','id pegawai','id_pegawai'],
    nama:     ['nama','name'],
    jabatan:  ['jabatan','posisi','position'],
    phone:    ['phone','no hp','nohp','telepon','no_hp','nomor_hp','nomor hp','no_telepon'],
    role:     ['role','peran'],
    password: ['password','pass','pw','kata sandi','kata_sandi']
  };
  const optionalAliases = { foto: ['foto','photo','foto_profil','foto profil'] };

  const map = {};
  Object.keys(aliases).forEach(function (key) {
    let idx = -1;
    for (let c = 0; c < headerRow.length; c++) {
      if (aliases[key].indexOf(normalize(headerRow[c])) !== -1) { idx = c; break; }
    }
    map[key] = idx;
  });
  Object.keys(optionalAliases).forEach(function (key) {
    let idx = -1;
    for (let c = 0; c < headerRow.length; c++) {
      if (optionalAliases[key].indexOf(normalize(headerRow[c])) !== -1) { idx = c; break; }
    }
    map[key] = idx;
  });

  const missing = Object.keys(aliases).filter(function (k) { return map[k] === -1; });
  if (missing.length > 0) {
    throw new Error("Header 'Users' tidak lengkap, kolom hilang: " + missing.join(', ') +
      ". Jalankan setupSheets().");
  }
  return map;
}

// ─────────────────────────────────────────────
//  PEMETAAN KOLOM ATTENDANCE (HEADER-BASED)
// ─────────────────────────────────────────────

function getAttendanceColumnMap() {
  const sheet     = getSheetSafe('Attendance');
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalize = function (s) { return String(s).trim().toLowerCase(); };

  const aliases = {
    id:       ['id','id_absen','idabsen'],
    userId:   ['userid','id_user','iduser'],
    nama:     ['nama','name'],
    tgl:      ['tgl','tanggal','date'],
    jamIn:    ['jamin','jam_masuk','jam masuk','checkin','check_in'],
    jamOut:   ['jamout','jam_keluar','jam keluar','checkout','check_out'],
    lat:      ['lat','lat_in','latin','latitude_in','latitude'],
    lng:      ['lng','lng_in','lngin','longitude_in','longitude'],
    latOut:   ['latout','lat_out','latitude_out'],
    lngOut:   ['lngout','lng_out','longitude_out'],
    status:   ['status'],
    photoIn:  ['photoin','foto masuk','foto_masuk','fotoin'],
    photoOut: ['photoout','foto keluar','foto_keluar','fotoout']
  };

  const map = {};
  Object.keys(aliases).forEach(function (key) {
    let idx = -1;
    for (let c = 0; c < headerRow.length; c++) {
      if (aliases[key].indexOf(normalize(headerRow[c])) !== -1) { idx = c; break; }
    }
    map[key] = idx;
  });
  return map;
}

// ─────────────────────────────────────────────
//  LOGIN
//  [FIX #3] Bandingkan hash password, bukan plaintext
// ─────────────────────────────────────────────

function checkLogin(phone, password) {
  try {
    // [FIX #5] Validasi input
    if (!phone || !password) return { success: false, message: 'Nomor HP dan password wajib diisi.' };

    const col        = getUserColumnMap();
    const sheet      = getSheetSafe('Users');
    const data       = sheet.getDataRange().getValues();
    const hashedInput = hashPassword(password);

    for (let i = 1; i < data.length; i++) {
      const storedPhone = String(data[i][col.phone]).trim();
      const storedPass  = String(data[i][col.password]).trim();

      if (storedPhone !== String(phone).trim()) continue;

      // Dukung password lama (plaintext) sekaligus hash baru
      const match = (storedPass === hashedInput) || (storedPass === password);
      if (match) {
        // Kalau masih plaintext, migrasi otomatis ke hash
        if (storedPass === password) {
          sheet.getRange(i + 1, col.password + 1).setValue(hashedInput);
        }
        return {
          success: true,
          user: {
            id:      data[i][col.id],
            nama:    data[i][col.nama],
            jabatan: data[i][col.jabatan],
            phone:   data[i][col.phone],
            role:    data[i][col.role]
          }
        };
      }
      // Phone cocok tapi password salah — hentikan pencarian
      return { success: false, message: 'Nomor HP atau Password salah!' };
    }
    return { success: false, message: 'Nomor HP atau Password salah!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  ABSENSI
//  [FIX #4] Normalisasi tanggal pakai formatTgl()
//  [FIX #5] Validasi input sebelum proses
// ─────────────────────────────────────────────

function submitAttendance(type, lat, lng, userId, selfieData) {
  try {
    // [FIX #5] Validasi input
    if (!type || !userId)   return { success: false, message: 'Parameter tidak lengkap.' };
    if (!selfieData)        return { success: false, message: 'Wajib mengambil foto selfie!' };
    if (isNaN(lat) || isNaN(lng)) return { success: false, message: 'Koordinat GPS tidak valid.' };

    const settings  = getSettings();
    const officeLat = normalizeCoordinate(settings.office_lat, 90);
    const officeLng = normalizeCoordinate(settings.office_lng, 180);
    const radius    = parseInt(settings.radius) || 100;
    const dist      = calculateDistance(lat, lng, officeLat, officeLng);

    if (dist > radius) {
      return { success: false, message: 'Gagal! Anda berada di luar radius kantor (' + Math.round(dist) + 'm).' };
    }

    const photoUrl = saveFileToDrive(selfieData, 'Selfie_' + userId + '_' + Date.now() + '.jpg');
    const sheet    = getSheetSafe('Attendance');
    const col      = getAttendanceColumnMap();
    const data     = sheet.getDataRange().getValues();

    // [FIX #4] Pakai formatTgl() agar konsisten
    const today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
    const now   = Utilities.formatDate(new Date(), 'GMT+7', 'HH:mm:ss');

    if (type === 'IN') {
      for (let i = 1; i < data.length; i++) {
        // [FIX #4] Normalisasi tanggal dari sheet
        if (String(data[i][col.userId]) == String(userId) && formatTgl(data[i][col.tgl]) == today) {
          return { success: false, message: 'Anda sudah absen masuk hari ini!' };
        }
      }

      const userList  = getNormalizedUsers();
      const userFound = userList.find(function(u) { return String(u[0]) == String(userId); });
      const namaUser  = userFound ? userFound[1] : '';

      const numCols = sheet.getLastColumn();
      const row     = new Array(numCols).fill('');
      row[col.id]       = generateId('Attendance'); // [FIX #2] ID aman
      row[col.userId]   = userId;
      if (col.nama !== undefined && col.nama >= 0) row[col.nama] = namaUser;
      row[col.tgl]      = today;
      row[col.jamIn]    = now;
      row[col.jamOut]   = '';
      row[col.lat]      = lat;
      row[col.lng]      = lng;
      row[col.latOut]   = '';
      row[col.lngOut]   = '';
      row[col.status]   = 'Hadir';
      row[col.photoIn]  = photoUrl;
      row[col.photoOut] = '';
      sheet.appendRow(row);
      return { success: true, message: 'Absen Masuk Berhasil!' };

    } else {
      for (let i = data.length - 1; i >= 1; i--) {
        // [FIX #4] Normalisasi tanggal dari sheet
        if (String(data[i][col.userId]) == String(userId) && formatTgl(data[i][col.tgl]) == today) {
          if (data[i][col.jamOut] !== '') return { success: false, message: 'Anda sudah absen keluar hari ini!' };
          sheet.getRange(i + 1, col.jamOut   + 1).setValue(now);
          sheet.getRange(i + 1, col.latOut   + 1).setValue(lat);
          sheet.getRange(i + 1, col.lngOut   + 1).setValue(lng);
          sheet.getRange(i + 1, col.photoOut + 1).setValue(photoUrl);
          return { success: true, message: 'Absen Keluar Berhasil!' };
        }
      }
      return { success: false, message: 'Anda belum melakukan absen masuk!' };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  IZIN / PERMIT
//  [FIX #5] Validasi input
// ─────────────────────────────────────────────

function submitPermit(permitData) {
  try {
    // [FIX #5] Validasi input
    if (!permitData || !permitData.userId) return { success: false, message: 'Data izin tidak valid.' };
    if (!permitData.date)                  return { success: false, message: 'Tanggal izin wajib diisi.' };
    if (!permitData.reason || !permitData.reason.trim()) return { success: false, message: 'Alasan izin wajib diisi.' };
    if (!permitData.fileData)              return { success: false, message: 'Lampiran foto wajib diunggah.' };

    let fileUrl = '';
    if (permitData.fileData && permitData.fileName) {
      const ext     = permitData.fileName.split('.').pop().toLowerCase();
      const mimeMap = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', pdf: 'application/pdf', gif: 'image/gif' };
      const mime    = mimeMap[ext] || 'application/octet-stream';
      fileUrl = saveFileToDrive('data:' + mime + ';base64,' + permitData.fileData, permitData.fileName);
    }

    const permitUserList  = getNormalizedUsers();
    const permitUserFound = permitUserList.find(function(u) { return String(u[0]) == String(permitData.userId); });
    const permitNama      = permitUserFound ? permitUserFound[1] : '';

    getSheetSafe('Permits').appendRow([
      generateId('Permits'), // [FIX #2] ID aman
      permitData.userId,
      permitNama,
      permitData.date,
      permitData.type,
      permitData.reason.trim(),
      fileUrl,
      'Menunggu'
    ]);
    return { success: true, message: 'Izin berhasil diajukan!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updatePermitStatus(id, status) {
  try {
    // [FIX #5] Validasi status
    const validStatuses = ['Disetujui', 'Ditolak', 'Menunggu'];
    if (!id)                              return { success: false, message: 'ID izin tidak valid.' };
    if (validStatuses.indexOf(status) === -1) return { success: false, message: 'Status tidak valid.' };

    const sheet     = getSheetSafe('Permits');
    const data      = sheet.getDataRange().getValues();
    const headerRow = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    const statusCol = headerRow.indexOf('status_app');
    if (statusCol === -1) return { success: false, message: "Kolom 'status_app' tidak ditemukan." };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(id)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(status);
        return { success: true };
      }
    }
    return { success: false, message: 'ID Izin tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  MANAJEMEN PEGAWAI
//  [FIX #2] generateId() aman
//  [FIX #3] Hash password
//  [FIX #5] Validasi input
// ─────────────────────────────────────────────

function addUser(userData) {
  try {
    // [FIX #5] Validasi wajib
    if (!userData.nama  || !userData.nama.trim())  return { success: false, message: 'Nama wajib diisi.' };
    if (!userData.phone || !userData.phone.trim())  return { success: false, message: 'Nomor HP wajib diisi.' };
    if (!userData.password)                         return { success: false, message: 'Password wajib diisi untuk pegawai baru.' };

    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');

    // Cek nomor HP duplikat
    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][col.phone]).trim() === String(userData.phone).trim()) {
        return { success: false, message: 'Nomor HP sudah terdaftar.' };
      }
    }

    const row = new Array(sheet.getLastColumn()).fill('');
    row[col.id]       = generateId('Users'); // [FIX #2]
    row[col.nama]     = userData.nama.trim();
    row[col.jabatan]  = (userData.jabatan || '').trim();
    row[col.phone]    = userData.phone.trim();
    row[col.role]     = userData.role || 'pegawai';
    row[col.password] = hashPassword(userData.password); // [FIX #3]
    sheet.appendRow(row);
    return { success: true, message: 'Pegawai berhasil ditambahkan!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function editUser(userData) {
  try {
    // [FIX #5] Validasi wajib
    if (!userData.nama  || !userData.nama.trim())  return { success: false, message: 'Nama wajib diisi.' };
    if (!userData.phone || !userData.phone.trim())  return { success: false, message: 'Nomor HP wajib diisi.' };

    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col.id]) == String(userData.id)) {
        const r = i + 1;
        sheet.getRange(r, col.nama    + 1).setValue(userData.nama.trim());
        sheet.getRange(r, col.jabatan + 1).setValue((userData.jabatan || '').trim());
        sheet.getRange(r, col.phone   + 1).setValue(userData.phone.trim());
        sheet.getRange(r, col.role    + 1).setValue(userData.role || 'pegawai');
        // [FIX #3] Hash jika password baru diisi
        if (userData.password) {
          sheet.getRange(r, col.password + 1).setValue(hashPassword(userData.password));
        }
        return { success: true, message: 'Data pegawai diperbarui!' };
      }
    }
    return { success: false, message: 'ID pegawai tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.toString(); }
  }
}

function deleteUser(id) {
  try {
    if (!id) return { success: false, message: 'ID tidak valid.' };
    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col.id]) == String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Pegawai dihapus!' };
      }
    }
    return { success: false, message: 'ID pegawai tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getUsers() {
  try { return getNormalizedUsers(); } catch (e) { return []; }
}

function getNormalizedUsers() {
  const col   = getUserColumnMap();
  const sheet = getSheetSafe('Users');
  // Password TIDAK dikirim ke frontend — hanya 5 kolom (tanpa index password)
  return sheet.getDataRange().getValues().slice(1).map(function (row) {
    return [row[col.id], row[col.nama], row[col.jabatan], row[col.phone], row[col.role]];
    //       [0]          [1]            [2]                [3]             [4]
    // Catatan: index [5] (password) sengaja dihilangkan demi keamanan
  });
}

// ─────────────────────────────────────────────
//  [FIX #1] DASHBOARD
//  allPermits SELALU diisi, tidak tergantung kondisi lain.
//  [FIX #4] Normalisasi tanggal dengan formatTgl()
// ─────────────────────────────────────────────

function getDashboardData(userId, role) {
  try {
    const att         = getSheetData('Attendance');
    const permits     = getSheetData('Permits');
    const users       = getNormalizedUsers();
    const rawSettings = getSettings();
    const settings    = Object.assign({}, rawSettings, {
      office_lat: normalizeCoordinate(rawSettings.office_lat, 90),
      office_lng: normalizeCoordinate(rawSettings.office_lng, 180)
    });

    // [FIX #4] Pakai formatTgl() untuk today
    const today       = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
    const attData     = att.slice(1);
    const permitsData = permits.slice(1);
    const attCol      = getAttendanceColumnMap();

    // Riwayat absensi user
    const historyNorm = attData
      .filter(function(r) { return String(r[attCol.userId]) == String(userId); })
      .slice(-10).reverse()
      .map(function(r) {
        return [
          r[attCol.id], r[attCol.userId],
          // [FIX #4] Normalisasi tanggal
          formatTgl(r[attCol.tgl]), r[attCol.jamIn], r[attCol.jamOut],
          r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
          r[attCol.status], r[attCol.photoIn], r[attCol.photoOut]
        ];
      });

    // Semua log absensi (untuk admin)
    const allAttNorm = attData.slice().reverse().map(function(r) {
      var found   = users.find(function(u) { return String(u[0]) == String(r[attCol.userId]); });
      var mapsUrl = 'https://www.google.com/maps?q=' + r[attCol.lat] + ',' + r[attCol.lng];
      var nama    = found ? found[1] : (attCol.nama >= 0 ? r[attCol.nama] : 'Anonim');
      return [
        r[attCol.id], r[attCol.userId],
        // [FIX #4] Normalisasi tanggal
        formatTgl(r[attCol.tgl]), r[attCol.jamIn], r[attCol.jamOut],
        r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
        r[attCol.status], r[attCol.photoIn], r[attCol.photoOut],
        mapsUrl, nama
      ];
    });

    // [FIX #4] Hitung hadir hari ini dengan normalisasi tanggal
    const hadirHariIni = attData.filter(function(r) {
      return formatTgl(r[attCol.tgl]) == today;
    }).length;

    return {
      stats: {
        totalPegawai: users.length,
        absenHariIni: hadirHariIni,
        izinPending:  permitsData.filter(function(r) { return r[7] == 'Menunggu'; }).length
      },
      history:       historyNorm,
      userPermits:   permitsData
                       .filter(function(r) { return String(r[1]) == String(userId); })
                       .slice(-10).reverse(),
      // [FIX #1] allPermits SELALU ada — tidak bersyarat
      allPermits:    permitsData.slice().reverse(),
      allAttendance: allAttNorm,
      allUsers:      users,
      settings:      settings
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  PENGATURAN
// ─────────────────────────────────────────────

function getSettings() {
  const d = getSheetSafe('Settings').getDataRange().getValues();
  const s = {};
  d.forEach(function (r) { s[String(r[0]).trim().toLowerCase()] = r[1]; });
  return s;
}

function getOfficeLocation() {
  try {
    var s = getSettings();
    return {
      lat: normalizeCoordinate(s.office_lat, 90)  || 0,
      lng: normalizeCoordinate(s.office_lng, 180) || 0,
      rad: parseInt(s.radius) || 100
    };
  } catch (e) {
    return { lat: 0, lng: 0, rad: 100, error: e.toString() };
  }
}

function updateOfficeSettings(lat, lng, radius) {
  try {
    // [FIX #5] Validasi koordinat
    const latNum = parseFloat(lat);
    const lngNum = parseFloat(lng);
    const radNum = parseInt(radius);
    if (isNaN(latNum) || Math.abs(latNum) > 90)   return { success: false, message: 'Latitude tidak valid (harus -90 s/d 90).' };
    if (isNaN(lngNum) || Math.abs(lngNum) > 180)  return { success: false, message: 'Longitude tidak valid (harus -180 s/d 180).' };
    if (isNaN(radNum) || radNum < 10 || radNum > 5000) return { success: false, message: 'Radius tidak valid (10–5000 meter).' };

    const s      = getSheetSafe('Settings');
    const data   = s.getDataRange().getValues();
    const wanted = { office_lat: latNum, office_lng: lngNum, radius: radNum };
    const found  = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      if (wanted.hasOwnProperty(key)) { s.getRange(i + 1, 2).setValue(wanted[key]); found[key] = true; }
    }
    Object.keys(wanted).forEach(function (key) { if (!found[key]) s.appendRow([key, wanted[key]]); });
    return { success: true, message: 'Lokasi kantor diperbarui!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  HELPER
// ─────────────────────────────────────────────

function saveFileToDrive(base64Data, fileName) {
  if (!base64Data || typeof base64Data !== 'string') throw new Error('saveFileToDrive: base64Data tidak valid');
  if (!base64Data.startsWith('data:')) throw new Error('saveFileToDrive: format harus data URL');

  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  const contentType = base64Data.substring(5, base64Data.indexOf(';'));
  const base64Only  = base64Data.split(',')[1];
  if (!contentType || !base64Only) throw new Error('saveFileToDrive: gagal parsing MIME/base64');

  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64Only), contentType, fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function normalizeCoordinate(value, maxAbs) {
  let n = parseFloat(value);
  if (isNaN(n) || n === 0) return 0;
  let g = 0;
  while (Math.abs(n) > maxAbs && g < 15) { n /= 10; g++; }
  if (Math.abs(n) > maxAbs) return 0;
  return n;
}

function getSheetData(n) {
  const sheet = getSS().getSheetByName(n);
  if (!sheet) return [[]];
  return sheet.getDataRange().getValues().map(function (row) {
    return row.map(function (cell) {
      if (cell instanceof Date) {
        // [FIX #4] Gunakan formatTgl untuk konsistensi
        return cell.getFullYear() === 1899
          ? Utilities.formatDate(cell, 'GMT+7', 'HH:mm:ss')
          : Utilities.formatDate(cell, 'GMT+7', 'yyyy-MM-dd');
      }
      return cell;
    });
  });
}

function getAttendanceLog() {
  try {
    var att    = getSheetData('Attendance');
    var users  = getNormalizedUsers();
    var attCol = getAttendanceColumnMap();
    return att.slice(1).reverse().map(function (r) {
      var found   = users.find(function (u) { return String(u[0]) == String(r[attCol.userId]); });
      var mapsUrl = 'https://www.google.com/maps?q=' + r[attCol.lat] + ',' + r[attCol.lng];
      var nama    = found ? found[1] : (attCol.nama >= 0 ? r[attCol.nama] : 'Anonim');
      return [
        r[attCol.id], r[attCol.userId],
        // [FIX #4] Normalisasi tanggal
        formatTgl(r[attCol.tgl]), r[attCol.jamIn], r[attCol.jamOut],
        r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
        r[attCol.status], r[attCol.photoIn], r[attCol.photoOut],
        mapsUrl, nama
      ];
    });
  } catch (e) { return []; }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R    = 6371e3;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a    = Math.sin(dLat / 2) ** 2 +
               Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
               Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
