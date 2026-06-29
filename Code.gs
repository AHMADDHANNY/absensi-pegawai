/** 
 * APLIKASI PRESENSI PEGAWAI - BACKEND CORE (FIXED v3)
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
//  GENERATE ID BERURUTAN
// ─────────────────────────────────────────────

/**
 * Buat ID angka berurutan otomatis berdasarkan jumlah baris di sheet.
 * Contoh: USR-1, USR-2, ATT-1, PRM-1
 */
function generateId(sheetName) {
  var sheet = getSS().getSheetByName(sheetName);
  if (!sheet) return 1;
  var lastRow = sheet.getLastRow();
  return lastRow <= 1 ? 1 : lastRow; // baris 1 = header, baris 2 = data pertama (id=1)
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
    ['office_lat',   0],
    ['office_lng',   0],
    ['radius',       100],
  ];
  defaults.forEach(function (d) {
    if (existingKeys.indexOf(d[0]) === -1) settingsSheet.appendRow(d);
  });

  var usersSheet = ss.getSheetByName('Users');
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([1,'Administrator','Admin','08000000000','admin123','admin']);
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
    .setTitle('Presensi Mobile Pro v3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────
//  FOLDER GOOGLE DRIVE
//  Ganti nama folder di bawah sesuai kebutuhan.
// ─────────────────────────────────────────────

var DRIVE_FOLDER_NAME = "FOTO ABSENSI";

// ─────────────────────────────────────────────
//  PENGATURAN NAMA FOLDER DRIVE
// ─────────────────────────────────────────────

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
    id:      ['id','id_absen','idabsen'],
    userId:  ['userid','id_user','iduser'],
    nama:    ['nama','name'],
    tgl:     ['tgl','tanggal','date'],
    jamIn:   ['jamin','jam_masuk','jam masuk','checkin','check_in'],
    jamOut:  ['jamout','jam_keluar','jam keluar','checkout','check_out'],
    lat:     ['lat','lat_in','latin','latitude_in','latitude'],
    lng:     ['lng','lng_in','lngin','longitude_in','longitude'],
    latOut:  ['latout','lat_out','latitude_out'],
    lngOut:  ['lngout','lng_out','longitude_out'],
    status:  ['status'],
    photoIn: ['photoin','foto masuk','foto_masuk','fotoin'],
    photoOut:['photoout','foto keluar','foto_keluar','fotoout']
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
// ─────────────────────────────────────────────

function checkLogin(phone, password) {
  try {
    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][col.phone].toString() == phone && data[i][col.password].toString() == password) {
        return {
          success: true,
          user: {
            id: data[i][col.id], nama: data[i][col.nama],
            jabatan: data[i][col.jabatan], phone: data[i][col.phone], role: data[i][col.role]
          }
        };
      }
    }
    return { success: false, message: "Nomor HP atau Password salah!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  ABSENSI
// ─────────────────────────────────────────────

function submitAttendance(type, lat, lng, userId, selfieData) {
  try {
    const settings  = getSettings();
    const officeLat = normalizeCoordinate(settings.office_lat, 90);
    const officeLng = normalizeCoordinate(settings.office_lng, 180);
    const radius    = parseInt(settings.radius) || 100;
    const dist      = calculateDistance(lat, lng, officeLat, officeLng);

    if (dist > radius) {
      return { success: false, message: "Gagal! Anda berada di luar radius kantor (" + Math.round(dist) + "m)." };
    }
    if (!selfieData) return { success: false, message: "Wajib mengambil foto selfie!" };

    const photoUrl = saveFileToDrive(selfieData, "Selfie_" + userId + "_" + Date.now() + ".jpg");
    const sheet    = getSheetSafe('Attendance');
    const col      = getAttendanceColumnMap();
    const data     = sheet.getDataRange().getValues();
    const today    = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    const now      = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");

    if (type === 'IN') {
      for (let i = 1; i < data.length; i++) {
        let d = data[i][col.tgl];
        if (d instanceof Date) d = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
        if (String(data[i][col.userId]) == String(userId) && d == today) {
          return { success: false, message: "Anda sudah absen masuk hari ini!" };
        }
      }
      // Ambil nama user dari sheet Users
      const userList  = getNormalizedUsers();
      const userFound = userList.find(function(u) { return String(u[0]) == String(userId); });
      const namaUser  = userFound ? userFound[1] : '';

      // Buat baris baru sesuai urutan kolom di sheet
      const numCols = sheet.getLastColumn();
      const row = new Array(numCols).fill('');
      row[col.id]      = generateId('Attendance');
      row[col.userId]  = userId;
      if (col.nama !== undefined && col.nama >= 0) row[col.nama] = namaUser;
      row[col.tgl]     = today;
      row[col.jamIn]   = now;
      row[col.jamOut]  = '';
      row[col.lat]     = lat;
      row[col.lng]     = lng;
      row[col.latOut]  = '';
      row[col.lngOut]  = '';
      row[col.status]  = 'Hadir';
      row[col.photoIn] = photoUrl;
      row[col.photoOut]= '';
      sheet.appendRow(row);
      return { success: true, message: "Absen Masuk Berhasil!" };

    } else {
      for (let i = data.length - 1; i >= 1; i--) {
        let d = data[i][col.tgl];
        if (d instanceof Date) d = Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
        if (String(data[i][col.userId]) == String(userId) && d == today) {
          if (data[i][col.jamOut] !== '') return { success: false, message: "Anda sudah absen keluar hari ini!" };
          sheet.getRange(i + 1, col.jamOut  + 1).setValue(now);
          sheet.getRange(i + 1, col.latOut  + 1).setValue(lat);
          sheet.getRange(i + 1, col.lngOut  + 1).setValue(lng);
          sheet.getRange(i + 1, col.photoOut+ 1).setValue(photoUrl);
          return { success: true, message: "Absen Keluar Berhasil!" };
        }
      }
      return { success: false, message: "Anda belum melakukan absen masuk!" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  IZIN / PERMIT
// ─────────────────────────────────────────────

function submitPermit(permitData) {
  try {
    let fileUrl = "";
    if (permitData.fileData && permitData.fileName) {
      const ext     = permitData.fileName.split('.').pop().toLowerCase();
      const mimeMap = { jpg:'image/jpeg', jpeg:'image/jpeg', png:'image/png', pdf:'application/pdf', gif:'image/gif' };
      const mime    = mimeMap[ext] || 'application/octet-stream';
      fileUrl = saveFileToDrive("data:" + mime + ";base64," + permitData.fileData, permitData.fileName);
    }
    const permitUserList  = getNormalizedUsers();
    const permitUserFound = permitUserList.find(function(u) { return String(u[0]) == String(permitData.userId); });
    const permitNama      = permitUserFound ? permitUserFound[1] : '';

    getSheetSafe('Permits').appendRow([
      generateId('Permits'), permitData.userId, permitNama, permitData.date,
      permitData.type, permitData.reason, fileUrl, "Menunggu"
    ]);
    return { success: true, message: "Izin berhasil diajukan!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function updatePermitStatus(id, status) {
  try {
    const sheet = getSheetSafe('Permits');
    const data  = sheet.getDataRange().getValues();
    // Cari index kolom status_app dari header
    const headerRow = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    const statusCol = headerRow.indexOf('status_app'); // kolom ke-8 (index 7)
    if (statusCol === -1) return { success: false, message: "Kolom 'status_app' tidak ditemukan di sheet Permits." };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) { sheet.getRange(i + 1, statusCol + 1).setValue(status); return { success: true }; }
    }
    return { success: false, message: "ID Izin tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────────
//  MANAJEMEN PEGAWAI
// ─────────────────────────────────────────────

function addUser(userData) {
  try {
    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const row   = new Array(sheet.getLastColumn()).fill('');
    row[col.id] = generateId('Users'); row[col.nama] = userData.nama;
    row[col.jabatan] = userData.jabatan; row[col.phone] = userData.phone;
    row[col.role] = userData.role; row[col.password] = userData.password;
    sheet.appendRow(row);
    return { success: true, message: "Pegawai berhasil ditambahkan!" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function editUser(userData) {
  try {
    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][col.id] == userData.id) {
        const r = i + 1;
        sheet.getRange(r, col.nama+1).setValue(userData.nama);
        sheet.getRange(r, col.jabatan+1).setValue(userData.jabatan);
        sheet.getRange(r, col.phone+1).setValue(userData.phone);
        sheet.getRange(r, col.role+1).setValue(userData.role);
        if (userData.password) sheet.getRange(r, col.password+1).setValue(userData.password);
        return { success: true, message: "Data pegawai diperbarui!" };
      }
    }
    return { success: false, message: "ID pegawai tidak ditemukan." };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function deleteUser(id) {
  try {
    const col   = getUserColumnMap();
    const sheet = getSheetSafe('Users');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][col.id] == id) { sheet.deleteRow(i + 1); return { success: true, message: "Pegawai dihapus!" }; }
    }
    return { success: false, message: "ID pegawai tidak ditemukan." };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getUsers() { try { return getNormalizedUsers(); } catch (e) { return []; } }

function getNormalizedUsers() {
  const col   = getUserColumnMap();
  const sheet = getSheetSafe('Users');
  return sheet.getDataRange().getValues().slice(1).map(function (row) {
    return [row[col.id], row[col.nama], row[col.jabatan], row[col.phone], row[col.role], row[col.password]];
  });
}

// ─────────────────────────────────────────────
//  DASHBOARD
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
    const today      = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    const attData    = att.slice(1);
    const permitsData = permits.slice(1);
    const attCol     = getAttendanceColumnMap();

    const historyNorm = attData
      .filter(function(r) { return String(r[attCol.userId]) == String(userId); })
      .slice(-10).reverse()
      .map(function(r) {
        return [
          r[attCol.id], r[attCol.userId],
          r[attCol.tgl], r[attCol.jamIn], r[attCol.jamOut],
          r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
          r[attCol.status], r[attCol.photoIn], r[attCol.photoOut]
        ];
      });

    const allAttNorm = attData.slice().reverse().map(function(r) {
      var found   = users.find(function(u) { return String(u[0]) == String(r[attCol.userId]); });
      var mapsUrl = 'https://www.google.com/maps?q=' + r[attCol.lat] + ',' + r[attCol.lng];
      var nama    = found ? found[1] : (attCol.nama >= 0 ? r[attCol.nama] : 'Anonim');
      return [
        r[attCol.id], r[attCol.userId],
        r[attCol.tgl], r[attCol.jamIn], r[attCol.jamOut],
        r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
        r[attCol.status], r[attCol.photoIn], r[attCol.photoOut],
        mapsUrl, nama
      ];
    });

    return {
      stats: {
        totalPegawai: users.length,
        absenHariIni: attData.filter(function(r) { return r[attCol.tgl] == today; }).length,
        izinPending:  permitsData.filter(function(r) { return r[7] == "Menunggu"; }).length
      },
      history:       historyNorm,
      userPermits:   permitsData.filter(function(r) { return String(r[1]) == String(userId); }).slice(-10).reverse(),
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
    return { lat: normalizeCoordinate(s.office_lat, 90) || 0, lng: normalizeCoordinate(s.office_lng, 180) || 0, rad: parseInt(s.radius) || 100 };
  } catch (e) { return { lat: 0, lng: 0, rad: 100, error: e.toString() }; }
}

function updateOfficeSettings(lat, lng, radius) {
  try {
    const s      = getSheetSafe('Settings');
    const data   = s.getDataRange().getValues();
    const wanted = { office_lat: lat, office_lng: lng, radius: radius };
    const found  = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      if (wanted.hasOwnProperty(key)) { s.getRange(i + 1, 2).setValue(wanted[key]); found[key] = true; }
    }
    Object.keys(wanted).forEach(function (key) { if (!found[key]) s.appendRow([key, wanted[key]]); });
    return { success: true, message: "Lokasi kantor diperbarui!" };
  } catch (e) { return { success: false, message: e.toString() }; }
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
  if (Math.abs(n) > maxAbs) return 0; // tetap tidak valid, kembalikan 0
  return n;
}

function getSheetData(n) {
  const sheet = getSS().getSheetByName(n);
  if (!sheet) return [[]];
  return sheet.getDataRange().getValues().map(function (row) {
    return row.map(function (cell) {
      if (cell instanceof Date) {
        return cell.getFullYear() === 1899
          ? Utilities.formatDate(cell, "GMT+7", "HH:mm:ss")
          : Utilities.formatDate(cell, "GMT+7", "yyyy-MM-dd");
      }
      return cell;
    });
  });
}

function getAttendanceLog() {
  try {
    var att   = getSheetData('Attendance');
    var users = getNormalizedUsers();
    var attCol = getAttendanceColumnMap();
    return att.slice(1).reverse().map(function (r) {
      var found   = users.find(function (u) { return String(u[0]) == String(r[attCol.userId]); });
      var mapsUrl = 'https://www.google.com/maps?q=' + r[attCol.lat] + ',' + r[attCol.lng];
      var nama    = found ? found[1] : (attCol.nama >= 0 ? r[attCol.nama] : 'Anonim');
      // Return dalam urutan tetap sesuai yang diharapkan frontend:
      // [0]id [1]userId [2]tanggal [3]jamIn [4]jamOut
      // [5]lat [6]lng [7]latOut [8]lngOut [9]status [10]photoIn [11]photoOut
      // [12]mapsUrl [13]userName
      return [
        r[attCol.id], r[attCol.userId],
        r[attCol.tgl], r[attCol.jamIn], r[attCol.jamOut],
        r[attCol.lat], r[attCol.lng], r[attCol.latOut], r[attCol.lngOut],
        r[attCol.status], r[attCol.photoIn], r[attCol.photoOut],
        mapsUrl, nama
      ];
    });
  } catch (e) { return []; }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3, dLat = (lat2-lat1)*Math.PI/180, dLon = (lon2-lon1)*Math.PI/180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*Math.sin(dLon/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}