// ==========================================
// FUNGSI SETUP OTORISASI GOOGLE DRIVE (PANCINGAN IZIN)
// ==========================================
function setupPermissions() {
  DriveApp.getRootFolder();
  DriveApp.createFile("dummy", "dummy"); // Memancing izin penuh untuk membuat file di Drive
  SpreadsheetApp.getActiveSpreadsheet();
  GmailApp.getInboxUnreadCount();
  DocumentApp.create("dummy"); // Memancing izin untuk Document API
  ScriptApp.getProjectTriggers(); // Memancing izin untuk fitur Auto-Delete (Penghancuran Otomatis)
}

// ==========================================
// FUNGSI UTAMA & KONFIGURASI TAMPILAN
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('public/index')
    .setTitle('Komite Keperawatan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// FUNGSI HELPER REAL-TIME STATUS BACKEND
// ==========================================
function calculateRealtimeStatusServer(masaBerlaku, originalStatus) {
  if (!masaBerlaku || masaBerlaku === "-" || String(masaBerlaku).trim() === "") return originalStatus || "-";
  if (String(masaBerlaku).trim() === "Seumur Hidup") return "Aktif";
  
  var d = new Date(masaBerlaku);
  if (isNaN(d.getTime())) return originalStatus || "-";
  
  var today = new Date();
  d.setHours(0,0,0,0);
  today.setHours(0,0,0,0);
  
  if (d >= today) {
    return "Aktif";
  } else {
    return "Habis Masa Berlaku";
  }
}

function calculateKredensialStatusServer(tanggalKredensial) {
  if (!tanggalKredensial || tanggalKredensial === "-" || String(tanggalKredensial).trim() === "") return "-";
  
  var d = new Date(tanggalKredensial);
  if (isNaN(d.getTime())) return "-";
  
  d.setFullYear(d.getFullYear() + 3); // Kredensial berlaku 3 tahun
  var today = new Date();
  
  d.setHours(0,0,0,0);
  today.setHours(0,0,0,0);
  
  if (today >= d) {
    return "Habis Masa Berlaku";
  } else {
    return "Aktif";
  }
}

// ==========================================
// FUNGSI HELPER OTORISASI & ROLE BACKEND (FASE 3)
// ==========================================
function _getUserRole(nip) {
  if (!nip) return "USER";
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = ss.getSheetByName("DATA KOMPER");
    if (!sheet) return "USER";
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return "USER";
    const headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    const idxNip = headers.indexOf("NIP");
    const idxRole = headers.indexOf("ROLE");
    if (idxNip === -1 || idxRole === -1) return "USER";
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxNip]).trim() === String(nip).trim()) {
        return String(data[i][idxRole] || "").trim().toUpperCase();
      }
    }
    return "USER";
  } catch(e) {
    return "USER";
  }
}

function _isAdminOrSuperAdmin(nip) {
  const role = _getUserRole(nip);
  return role === "ADMIN" || role === "SUPER ADMIN";
}

function _isSuperAdmin(nip) {
  const role = _getUserRole(nip);
  return role === "SUPER ADMIN";
}

function _canEditNip(targetNip, actorNip) {
  if (!targetNip) return false;
  if (!actorNip) return false;
  if (String(targetNip).trim() === String(actorNip).trim()) return true;
  return _isAdminOrSuperAdmin(actorNip);
}

// ==========================================
// FUNGSI LOGIN & LUPA PASSWORD
// ==========================================
function prosesLogin(nip, password) {
  try {
    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
    if(!sheet) return { success: false, message: "Sheet DATA KOMPER tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    
    const nipIdx = headers.indexOf("NIP");
    const passIdx = headers.indexOf("PASSWORD");
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIdx]).trim() === String(nip).trim() && String(data[i][passIdx]) === String(password)) {
        let userObj = {};
        for (let j = 0; j < headers.length; j++) {
          let val = data[i][j];
          if (val instanceof Date) {
            val = val.toString();
          }
          userObj[headers[j]] = val !== undefined ? val : "";
        }
      
      let calculatedKredensialStatus = calculateKredensialStatusServer(userObj["TANGGAL KREDENSIAL"]);
      if (calculatedKredensialStatus !== "-" && calculatedKredensialStatus !== userObj["STATUS KREDENSIAL"]) {
          userObj["STATUS KREDENSIAL"] = calculatedKredensialStatus;
          let statusKredIdx = headers.indexOf("STATUS KREDENSIAL");
          if (statusKredIdx !== -1) {
              sheet.getRange(i + 1, statusKredIdx + 1).setValue(calculatedKredensialStatus);
          }
      }
      
      let calculatedStrStatus = calculateRealtimeStatusServer(userObj["MASA BERLAKU STR AKTIF"], userObj["STATUS STR"]);
      if (calculatedStrStatus !== "-" && calculatedStrStatus !== userObj["STATUS STR"]) {
          userObj["STATUS STR"] = calculatedStrStatus;
          let statusStrIdx = headers.indexOf("STATUS STR");
          if (statusStrIdx !== -1) {
              sheet.getRange(i + 1, statusStrIdx + 1).setValue(calculatedStrStatus);
          }
      }
      
      let calculatedSippStatus = calculateRealtimeStatusServer(userObj["MASA BERLAKU SIPP SIPB AKTIF"], userObj["STATUS SIPP SIPB"]);
      if (calculatedSippStatus !== "-" && calculatedSippStatus !== userObj["STATUS SIPP SIPB"]) {
          userObj["STATUS SIPP SIPB"] = calculatedSippStatus;
          let statusSippIdx = headers.indexOf("STATUS SIPP SIPB");
          if (statusSippIdx !== -1) {
              sheet.getRange(i + 1, statusSippIdx + 1).setValue(calculatedSippStatus);
          }
      }
      
      let calculatedBtclsStatus = calculateRealtimeStatusServer(userObj["MASA BERLAKU BTCLS"], userObj["STATUS BTCLS"]);
      if (calculatedBtclsStatus !== "-" && calculatedBtclsStatus !== userObj["STATUS BTCLS"]) {
          userObj["STATUS BTCLS"] = calculatedBtclsStatus;
          let statusBtclsIdx = headers.indexOf("STATUS BTCLS");
          if (statusBtclsIdx !== -1) {
              sheet.getRange(i + 1, statusBtclsIdx + 1).setValue(calculatedBtclsStatus);
          }
      }
      
        let calculatedMuStatus = calculateRealtimeStatusServer(userObj["MASA BERLAKU PELATIHAN MIDWIFE UPDATE"], userObj["STATUS PELATIHAN MU"]);
        if (calculatedMuStatus !== "-" && calculatedMuStatus !== userObj["STATUS PELATIHAN MU"]) {
            userObj["STATUS PELATIHAN MU"] = calculatedMuStatus;
            let statusMuIdx = headers.indexOf("STATUS PELATIHAN MU");
            if (statusMuIdx !== -1) {
                sheet.getRange(i + 1, statusMuIdx + 1).setValue(calculatedMuStatus);
            }
        }
        
        // Ambil otorisasi menu
        let userAccess = getUserAccess(userObj["NIP"]);
        
        return {
          success: true,
          data: {
          nip: userObj["NIP"],
          password: userObj["PASSWORD"],
          nama: userObj["NAMA"],
          role: userObj["ROLE"],
          foto: userObj["FOTO"],
          ktp: userObj["NOMOR KTP"],
          tempat_lahir: userObj["TEMPAT LAHIR"],
          tanggal_lahir: userObj["TANGGAL LAHIR"],
          jk: userObj["JENIS KELAMIN"],
          agama: userObj["AGAMA"],
          alamat: userObj["ALAMAT"],
          rt: userObj["RT"],
          rw: userObj["RW"],
          provinsi: userObj["PROVINSI"],
          kabupaten: userObj["KABUPATEN"],
          kecamatan: userObj["KECAMATAN"],
          kelurahan: userObj["KELURAHAN"],
          hp: userObj["NOMOR HANDPHONE"],
          email: userObj["ALAMAT EMAIL"],
          npwp: userObj["NOMOR NPWP"],
          link_ktp: userObj["LINK KTP"],
          
          status_kepegawaian: userObj["STATUS KEPEGAWAIAN"],
          nomor_karpeg: userObj["NOMOR KARPEG"],
          golongan_terakhir: userObj["GOLONGAN TERAKHIR"],
          tmt_golongan_terakhir: userObj["TMT GOLONGAN TERAKHIR"],
          masa_kerja_sesuai_sk: userObj["MASA KERJA SESUAI SK"],
          gaji_pokok: userObj["GAJI POKOK"],
          link_sk_terakhir: userObj["LINK SK TERAKHIR"],
          jabatan_fungsional: userObj["JABATAN FUNGSIONAL"],
          tmt_jabatan_fungsional: userObj["TMT JABATAN FUNGSIONAL"],
          link_jabatan_fungsional: userObj["LINK JABATAN FUNGSIONAL"],
          sk_kgb: userObj["SK KGB"],
          tmt_sk_kgb: userObj["TMT SK KGB"],
          gaji_pokok_kgb: userObj["GAJI POKOK KGB"],
          link_sk_kgb: userObj["LINK SK KGB"],
          
          ruangan_unit: userObj["RUANGAN UNIT"],
          kelompok: userObj["KELOMPOK"],
          jabatan: userObj["JABATAN"],
          jenjang_karir: userObj["JENJANG KARIR"],
          tanggal_kredensial: userObj["TANGGAL KREDENSIAL"],
          link_sertifikat_jenjang_karir: userObj["LINK SERTIFIKAT JENJANG KARIR"],
          status_kredensial: userObj["STATUS KREDENSIAL"],
          
          institusi_pendidikan_terakhir: userObj["INSTITUSI PENDIDIKAN TERAKHIR"],
          jurusan: userObj["JURUSAN"],
          tahun_lulus_pendidikan_terakhir: userObj["TAHUN LULUS PENDIDIKAN TERAKHIR"],
          jenjang_pendidikan_terakhir: userObj["JENJANG PENDIDIKAN TERAKHIR"],
          link_ijazah: userObj["LINK IJAZAH"],
          
          nomor_str_aktif: userObj["NOMOR STR AKTIF"],
          tanggal_terbit_str_aktif: userObj["TANGGAL TERBIT STR AKTIF"],
          masa_berlaku_str_aktif: userObj["MASA BERLAKU STR AKTIF"],
          link_str: userObj["LINK STR"],
          status_str: calculateRealtimeStatusServer(userObj["MASA BERLAKU STR AKTIF"], userObj["STATUS STR"]),
          
          nomor_sipp_sipb_aktif: userObj["NOMOR SIPP SIPB AKTIF"],
          tanggal_terbit_sipp_sipb_aktif: userObj["TANGGAL TERBIT SIPP SIPB AKTIF"],
          masa_berlaku_sipp_sipb_aktif: userObj["MASA BERLAKU SIPP SIPB AKTIF"],
          link_sipp_sipb: userObj["LINK SIPP SIPB"],
          status_sipp_sipb: calculateRealtimeStatusServer(userObj["MASA BERLAKU SIPP SIPB AKTIF"], userObj["STATUS SIPP SIPB"]),
          
          nomor_btcls: userObj["NOMOR BTCLS"],
          tanggal_terbit_btcls: userObj["TANGGAL TERBIT BTCLS"],
          masa_berlaku_btcls: userObj["MASA BERLAKU BTCLS"],
          link_btcls: userObj["LINK BTCLS"],
          status_btcls: calculateRealtimeStatusServer(userObj["MASA BERLAKU BTCLS"], userObj["STATUS BTCLS"]),
          
          nomor_pelatihan_kamar_operasi: userObj["NOMOR PELATIHAN KAMAR OPERASI"],
          tanggal_terbit_pelatihan_kamar_operasi: userObj["TANGGAL TERBIT PELATIHAN KAMAR OPERASI"],
          masa_berlaku_pelatihan_kamar_operasi: userObj["MASA BERLAKU PELATIHAN KAMAR OPERASI"],
          link_pelatihan_kamar_operasi: userObj["LINK PELATIHAN KAMAR OPERASI"],
          status_pelatihan_kamar_operasi: calculateRealtimeStatusServer(userObj["MASA BERLAKU PELATIHAN KAMAR OPERASI"], userObj["STATUS PELATIHAN KAMAR OPERASI"]),
          
          nomor_pelatihan_intensif: userObj["NOMOR PELATIHAN INTENSIF"],
          tanggal_terbit_pelatihan_intensif: userObj["TANGGAL TERBIT PELATIHAN INTENSIF"],
          masa_berlaku_pelatihan_intensif: userObj["MASA BERLAKU PELATIHAN INTENSIF"],
          link_pelatihan_intensif: userObj["LINK PELATIHAN INTENSIF"],
          status_pelatihan_intensif: calculateRealtimeStatusServer(userObj["MASA BERLAKU PELATIHAN INTENSIF"], userObj["STATUS PELATIHAN INTENSIF"]),
          
          nomor_pelatihan_hd: userObj["NOMOR PELATIHAN HD"],
          tanggal_terbit_pelatihan_hd: userObj["TANGGAL TERBIT PELATIHAN HD"],
          masa_berlaku_pelatihan_hd: userObj["MASA BERLAKU PELATIHAN HD"],
          link_pelatihan_hd: userObj["LINK PELATIHAN HD"],
          status_pelatihan_hd: calculateRealtimeStatusServer(userObj["MASA BERLAKU PELATIHAN HD"], userObj["STATUS PELATIHAN HD"]),
          
          nomor_pelatihan_midwife_update: userObj["NOMOR PELATIHAN MIDWIFE UPDATE"],
          tanggal_terbit_pelatihan_midwife_update: userObj["TANGGAL TERBIT PELATIHAN MIDWIFE UPDATE"],
          masa_berlaku_pelatihan_midwife_update: userObj["MASA BERLAKU PELATIHAN MIDWIFE UPDATE"],
          link_pelatihan_mu: userObj["LINK PELATIHAN MU"],
            status_pelatihan_mu: calculateRealtimeStatusServer(userObj["MASA BERLAKU PELATIHAN MIDWIFE UPDATE"], userObj["STATUS PELATIHAN MU"]),
            
            verifikasi: userObj["VERIFIKASI"],
            menu_akses: userAccess.menus
          }
        };
    }
  }
  return { success: false, message: "NIP atau Password salah." };
  } catch (e) {
    return { success: false, message: "Sistem sibuk atau gagal memuat data: " + e.message };
  }
}

function prosesLupaPassword(nip, inputEmail) {
  const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
  if(!sheet) return { success: false, message: "Sistem belum siap. Sheet tidak ditemukan." };

  const data = sheet.getDataRange().getValues();
  const headersRaw = data[0];
  const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
  
  const nipIdx = headers.indexOf("NIP");
  const emailIdx = headers.indexOf("ALAMAT EMAIL");
  const passIdx = headers.indexOf("PASSWORD");
  const namaIdx = headers.indexOf("NAMA");
  
  if (nipIdx === -1 || emailIdx === -1 || passIdx === -1) {
    return { success: false, message: "Kolom NIP, ALAMAT EMAIL, atau PASSWORD tidak ditemukan di database." };
  }
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nipIdx]).trim() === String(nip).trim()) { 
      
      let emailRaw = String(data[i][emailIdx] || "").toLowerCase();
      const email = emailRaw.split(',')[0].trim();
      const nama = data[i][namaIdx];
      
      if (!inputEmail || String(inputEmail).trim().toLowerCase() !== email) {
        return { success: false, message: "Kombinasi NIP dan Email tidak cocok dengan database kami." };
      }
      
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!email || email === "-" || !emailRegex.test(email)) {
        return { success: false, message: "Email (" + (email || "Kosong") + ") tidak valid. Silakan Hubungi Admin Komite Keperawatan." };
      }
      
      const newPassword = Math.floor(100000 + Math.random() * 900000).toString();
      
        try {
          const subject = "Pemulihan Password Komite Keperawatan RSUD Subang";
          const body = "Halo " + nama + ",\n\nBerikut adalah password baru Anda untuk masuk ke aplikasi Komite Keperawatan RSUD Subang:\n\nPassword Baru: " + newPassword + "\n\nHarap simpan password ini dengan baik.\n\nTerima Kasih,\nAdmin Komite Keperawatan RSUD Subang";
          
          // Kirim email dulu, jika gagal akan masuk ke catch dan password belum diubah
          GmailApp.sendEmail(email, subject, body);
          
          sheet.getRange(i + 1, passIdx + 1).setValue(newPassword);
          SpreadsheetApp.flush();
  
          return { success: true, message: "Password berhasil di-reset dan dikirim ke email." };
      } catch(e) {
        return { success: false, message: "Gagal mengirim email. Detail Error: " + e.message };
      }
    }
  }
  return { success: false, message: "NIP tidak ditemukan dalam sistem." };
}

// ==========================================
// FUNGSI PENGAMBILAN REFERENSI (DROPDOWN/SLIDESHOW)
// ==========================================
function getSlideshowImages() {
  const images = [];
  
  // 1. Ambil dari Google Drive Folder
  try {
    const folderId = '1J3R4trtlJ9QaabMWYLC-c12ieYVEeUTX';
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      if (mimeType.indexOf('image/') !== -1) {
        images.push(file.getUrl());
      }
    }
  } catch(e) {
  }
  
  // 2. Ambil dari Sheet SLIDESHOW
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = ss.getSheetByName("SLIDESHOW");
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1]) { // Kolom URL
          images.push(data[i][1]);
        }
      }
    }
  } catch(e) {
  }
  
  return images;
}

function getDropdownAlamat(jenis, parentName) {
  const cache = CacheService.getScriptCache();
  const safeParent = String(parentName || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "_");
  const cacheKey = "ADDR_" + jenis + "_" + safeParent;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = ss.getSheetByName("ALAMAT");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    
    const provIdx = headers.indexOf("PROVINSI");
    const kabIdx = headers.indexOf("KABUPATEN");
    const kecIdx = headers.indexOf("KECAMATAN");
    let kelIdx = headers.indexOf("KELURAHAN");
    if (kelIdx === -1) kelIdx = headers.indexOf("KELUARAHAN"); 
    
    let resultSet = {};
    let parentStr = String(parentName).trim().toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let valProv = provIdx !== -1 && row[provIdx] ? String(row[provIdx]).trim() : "";
      let valKab = kabIdx !== -1 && row[kabIdx] ? String(row[kabIdx]).trim() : "";
      let valKec = kecIdx !== -1 && row[kecIdx] ? String(row[kecIdx]).trim() : "";
      let valKel = kelIdx !== -1 && row[kelIdx] ? String(row[kelIdx]).trim() : "";
      
      if (jenis === 'PROVINSI') {
        if (valProv) resultSet[valProv] = true;
      } else if (jenis === 'KABUPATEN') {
        if (valProv.toUpperCase() === parentStr && valKab) resultSet[valKab] = true;
      } else if (jenis === 'KECAMATAN') {
        if (valKab.toUpperCase() === parentStr && valKec) resultSet[valKec] = true;
      } else if (jenis === 'KELURAHAN') {
        if (valKec.toUpperCase() === parentStr && valKel) resultSet[valKel] = true;
      }
    }
    
    const sortedResult = Object.keys(resultSet).sort();
    try {
      const jsonStr = JSON.stringify(sortedResult);
      if (jsonStr.length < 90000) {
        cache.put(cacheKey, jsonStr, 21600); // Cache 6 Jam
      }
    } catch(cErr) {}

    return sortedResult;
  } catch(e) {
    return [];
  }
}

function getDropdownUnit() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("UNIT_DROPDOWN");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("UNIT");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let units = [];
  for (let i = 1; i < data.length; i++) { 
    if(data[i][0]){
       units.push({
         ruangan: data[i][0] || "",
         kelompok: data[i][1] || ""
       });
    }
  }
  try {
    cache.put("UNIT_DROPDOWN", JSON.stringify(units), 21600);
  } catch(e) {}

  return units;
}

// ==========================================
// FUNGSI DROPDOWN JABATAN FUNGSIONAL
// ==========================================
function getDropdownJabfung() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("JABFUNG_DROPDOWN");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  try {
    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("JABFUNG");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getDisplayValues();
    let jabfung = [];
    
    for (let i = 1; i < data.length; i++) { 
      if(data[i][0]) {
         jabfung.push(String(data[i][0]).trim());
      }
    }
    try {
      cache.put("JABFUNG_DROPDOWN", JSON.stringify(jabfung), 21600);
    } catch(e) {}

    return jabfung;
  } catch (error) {
    return [];
  }
}

// ==========================================
// FUNGSI UPDATE DATA KE SPREADSHEET (MENU PROFIL)
// ==========================================
function updateDataByNip(nip, updateMap, actorNip) {
  const effectiveActor = actorNip || nip;
  if (!_canEditNip(nip, effectiveActor)) {
    return { success: false, message: "Akses ditolak: Anda tidak memiliki wewenang untuk mengubah data pegawai ini." };
  }

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return { success: false, message: "Server sedang sibuk memproses antrean. Silakan coba sesaat lagi." };
    }

    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
    if(!sheet) return { success: false, message: "Sheet DATA KOMPER tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    const nipIndex = headers.indexOf("NIP");
    
    if (nipIndex === -1) return { success: false, message: "Kolom NIP tidak ditemukan di Sheet." };

    const isSA = _isSuperAdmin(effectiveActor);
    const isAdmin = _isAdminOrSuperAdmin(effectiveActor);

    // Keamanan: Cek field sensitif (ROLE hanya boleh Super Admin, VERIFIKASI hanya boleh Admin/Super Admin)
    let safeUpdateMap = {};
    for (let key in updateMap) {
      let upperKey = String(key).trim().toUpperCase();
      if (upperKey === "ROLE" && !isSA) continue;
      if (upperKey === "VERIFIKASI" && !isAdmin) continue;
      safeUpdateMap[upperKey] = updateMap[key];
    }
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIndex]).trim() === String(nip).trim()) {
        let rowToUpdate = i + 1;
        let rowData = data[i].slice();
        
        for (let key in safeUpdateMap) {
          let colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            rowData[colIndex] = safeUpdateMap[key];
          }
        }
        
        // Optimasi (Fase 4): Batch write 1 baris sekaligus via setValues
        sheet.getRange(rowToUpdate, 1, 1, headers.length).setValues([rowData]);
        return { success: true, message: "Data berhasil diperbarui." };
      }
    }
    return { success: false, message: "Data NIP tidak ditemukan." };
  } catch (e) {
    return { success: false, message: "Error sistem: " + e.message };
  } finally {
    try { lock.releaseLock(); } catch(lErr) {}
  }
}

function updateBiodata(d, actorNip) {
  if (!d || !d.nip) {
    return { success: false, message: "Data tidak valid atau gagal diterima server." };
  }
  const effectiveActor = actorNip || d.actorNip || d.nip;
  
  return updateDataByNip(d.nip, {
    "PASSWORD": d.password,
    "NAMA": d.nama,
    "NOMOR KTP": d.ktp,
    "TEMPAT LAHIR": d.tempat_lahir,
    "TANGGAL LAHIR": d.tanggal_lahir,
    "JENIS KELAMIN": d.jk,
    "AGAMA": d.agama,
    "ALAMAT": d.alamat,
    "RT": d.rt,
    "RW": d.rw,
    "PROVINSI": d.provinsi_name,
    "KABUPATEN": d.kabupaten_name,
    "KECAMATAN": d.kecamatan_name,
    "KELURAHAN": d.kelurahan_name,
    "NOMOR HANDPHONE": d.hp,
    "ALAMAT EMAIL": d.email,
    "NOMOR NPWP": d.npwp ? "'" + String(d.npwp) : "",
    "LINK KTP": d.link_ktp
  }, effectiveActor);
}

function updateKepegawaian(d, actorNip) {
  if (!d || !d.nip) return { success: false, message: "Data tidak valid." };
  const effectiveActor = actorNip || d.actorNip || d.nip;
  return updateDataByNip(d.nip, {
    "STATUS KEPEGAWAIAN": d.status_kepegawaian,
    "NOMOR KARPEG": d.nomor_karpeg,
    "GOLONGAN TERAKHIR": d.golongan_terakhir,
    "TMT GOLONGAN TERAKHIR": d.tmt_golongan_terakhir,
    "MASA KERJA SESUAI SK": d.masa_kerja_sesuai_sk,
    "GAJI POKOK": d.gaji_pokok,
    "LINK SK TERAKHIR": d.link_sk_terakhir,
    "JABATAN FUNGSIONAL": d.jabatan_fungsional,
    "TMT JABATAN FUNGSIONAL": d.tmt_jabatan_fungsional,
    "LINK JABATAN FUNGSIONAL": d.link_jabatan_fungsional,
    "SK KGB": d.sk_kgb,
    "TMT SK KGB": d.tmt_sk_kgb,
    "GAJI POKOK KGB": d.gaji_pokok_kgb,
    "LINK SK KGB": d.link_sk_kgb
  }, effectiveActor);
}

function updateUnitRuangan(d, actorNip) {
  if (!d || !d.nip) return { success: false, message: "Data tidak valid." };
  const effectiveActor = actorNip || d.actorNip || d.nip;
  return updateDataByNip(d.nip, {
    "RUANGAN UNIT": d.ruangan_unit,
    "KELOMPOK": d.kelompok,
    "JABATAN": d.jabatan,
    "JENJANG KARIR": d.jenjang_karir,
    "TANGGAL KREDENSIAL": d.tanggal_kredensial,
    "LINK SERTIFIKAT JENJANG KARIR": d.link_sertifikat_jenjang_karir,
    "STATUS KREDENSIAL": d.status_kredensial
  }, effectiveActor);
}

function updatePendidikan(d, actorNip) {
  if (!d || !d.nip) return { success: false, message: "Data tidak valid." };
  const effectiveActor = actorNip || d.actorNip || d.nip;
  return updateDataByNip(d.nip, {
    "INSTITUSI PENDIDIKAN TERAKHIR": d.institusi_pendidikan_terakhir,
    "JURUSAN": d.jurusan,
    "TAHUN LULUS PENDIDIKAN TERAKHIR": d.tahun_lulus_pendidikan_terakhir,
    "JENJANG PENDIDIKAN TERAKHIR": d.jenjang_pendidikan_terakhir,
    "LINK IJAZAH": d.link_ijazah
  }, effectiveActor);
}

function updateKompetensi(d, actorNip) {
  if (!d || !d.nip) return { success: false, message: "Data tidak valid." };
  const effectiveActor = actorNip || d.actorNip || d.nip;
  return updateDataByNip(d.nip, {
    "NOMOR STR AKTIF": d.nomor_str, 
    "TANGGAL TERBIT STR AKTIF": d.tgl_terbit_str, 
    "MASA BERLAKU STR AKTIF": d.masa_berlaku_str, 
    "STATUS STR": d.status_str, 
    "LINK STR": d.link_str,
    
    "NOMOR SIPP SIPB AKTIF": d.nomor_sipp, 
    "TANGGAL TERBIT SIPP SIPB AKTIF": d.tgl_terbit_sipp, 
    "MASA BERLAKU SIPP SIPB AKTIF": d.masa_berlaku_sipp, 
    "STATUS SIPP SIPB": d.status_sipp, 
    "LINK SIPP SIPB": d.link_sipp,
    
    "NOMOR BTCLS": d.nomor_btcls, 
    "TANGGAL TERBIT BTCLS": d.tgl_terbit_btcls, 
    "MASA BERLAKU BTCLS": d.masa_berlaku_btcls, 
    "STATUS BTCLS": d.status_btcls, 
    "LINK BTCLS": d.link_btcls,
    
    "NOMOR PELATIHAN KAMAR OPERASI": d.nomor_kamar_operasi, 
    "TANGGAL TERBIT PELATIHAN KAMAR OPERASI": d.tgl_terbit_kamar_operasi, 
    "MASA BERLAKU PELATIHAN KAMAR OPERASI": d.masa_berlaku_kamar_operasi, 
    "STATUS PELATIHAN KAMAR OPERASI": d.status_kamar_operasi, 
    "LINK PELATIHAN KAMAR OPERASI": d.link_kamar_operasi,
    
    "NOMOR PELATIHAN INTENSIF": d.nomor_intensif, 
    "TANGGAL TERBIT PELATIHAN INTENSIF": d.tgl_terbit_intensif, 
    "MASA BERLAKU PELATIHAN INTENSIF": d.masa_berlaku_intensif, 
    "STATUS PELATIHAN INTENSIF": d.status_intensif, 
    "LINK PELATIHAN INTENSIF": d.link_intensif,
    
    "NOMOR PELATIHAN HD": d.nomor_hd, 
    "TANGGAL TERBIT PELATIHAN HD": d.tgl_terbit_hd, 
    "MASA BERLAKU PELATIHAN HD": d.masa_berlaku_hd, 
    "STATUS PELATIHAN HD": d.status_hd, 
    "LINK PELATIHAN HD": d.link_hd,
    
    "NOMOR PELATIHAN MIDWIFE UPDATE": d.nomor_mu, 
    "TANGGAL TERBIT PELATIHAN MIDWIFE UPDATE": d.tgl_terbit_mu, 
    "MASA BERLAKU PELATIHAN MIDWIFE UPDATE": d.masa_berlaku_mu, 
    "STATUS PELATIHAN MU": d.status_mu, 
    "LINK PELATIHAN MU": d.link_mu
  }, effectiveActor);
}

// ==========================================
// FUNGSI UPLOAD FILE KE GOOGLE DRIVE
// ==========================================
function genericUpload(base64Data, nip, folderId, prefixName, colName, actorNip) {
  try {
    const effectiveActor = actorNip || nip;
    if (!_canEditNip(nip, effectiveActor)) {
      return { success: false, message: "Akses ditolak untuk mengunggah file atas NIP ini." };
    }

    let type = 'application/octet-stream';
    let base64String = base64Data;
    let ext = '';
    
    if (base64Data.indexOf(',') !== -1) {
      const splitBase = base64Data.split(',');
      type = splitBase[0].split(';')[0].replace('data:', '');
      base64String = splitBase[1];
    }
    
    if (type.indexOf('pdf') !== -1) ext = '.pdf';
    else if (type.indexOf('png') !== -1) ext = '.png';
    else if (type.indexOf('jpeg') !== -1 || type.indexOf('jpg') !== -1) ext = '.jpg';
    
    const byteCharacters = Utilities.base64Decode(base64String);
    const fileName = nip + "_" + prefixName + "_" + new Date().getTime() + ext;
    const blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    let folder;
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    const file = folder.createFile(blob);
    
    // Perbaikan: Hindari error jika domain memblokir akses file publik
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shareErr) {
      // Abaikan jika ditolak, file akan tetap tersimpan
    }
    
    const url = file.getUrl();

    if (colName) {
      let updateObj = {};
      updateObj[colName] = url;
      updateDataByNip(nip, updateObj, effectiveActor);
    }
    
    return { success: true, url: url };
  } catch (e) {
    return { success: false, message: "Gagal upload " + prefixName + ": " + e.message };
  }
}

function uploadFoto(b, nip, actorNip) { return genericUpload(b, nip, '1kDiRXeHHy2RZPN8UYWNBzHdECF-yOo-M', 'Foto', 'FOTO', actorNip); }
function uploadKtpToDrive(b, nip, actorNip) { return genericUpload(b, nip, '18pPTIOIAIyc61iWck46TpRs8AawV53Nm', 'KTP', null, actorNip); }

function uploadSkTerakhirToDrive(b, nip, actorNip) { return genericUpload(b, nip, '15jSTi3r-lxmVFVSd6hhhcNY9iTSKxpz6', 'SK_Terakhir', null, actorNip); }
function uploadJabatanFungsionalToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1vQOzZT0qC0OCO5UDlZAF2RGUfrvnt-6B', 'Jabatan_Fungsional', null, actorNip); }
function uploadSkKgbToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1Wm9B7NTkmud-W4Fcr1fQlumWz5xGfL1P', 'SK_KGB', null, actorNip); }

function uploadSertifikatJenjangKarirToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1_DdPyavY-_lJuf8iDmUvMFw3m8751ASY', 'Sertifikat_Karir', null, actorNip); }
function uploadIjazahToDrive(b, nip, actorNip) { return genericUpload(b, nip, '11MrcbHnlU-f4kZ2RjHzip03LiHwqhVtq', 'Ijazah', null, actorNip); }

function uploadStrToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1kKMCh9HE-hZui1I3PS4eZIt9eks-oVNb', 'STR', null, actorNip); }
function uploadSippToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1jWp4p3KXlnWapEOSm34k8qaimwIYMunL', 'SIPP', null, actorNip); }
function uploadBtclsToDrive(b, nip, actorNip) { return genericUpload(b, nip, '12otLna1I5EieIcoyWe_fpaca9TQo5LDB', 'BTCLS', null, actorNip); }
function uploadKamarOperasiToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1vCGRuQSvMEcczjOHXUkDNKNoe-y6CBVe', 'Kamar_Operasi', null, actorNip); }
function uploadIntensifToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1xAVXvwNCWak7V1t4nl7q6lfz-U9y-vEt', 'Intensif', null, actorNip); }
function uploadHdToDrive(b, nip, actorNip) { return genericUpload(b, nip, '11iY_iZ6WlHJ9oXESDQXKWpe8q2MSSel3', 'HD', null, actorNip); }
function uploadMuToDrive(b, nip, actorNip) { return genericUpload(b, nip, '1sNThKdiOJPWVSSE8y06cSuY9wQnQjV6c', 'MU', null, actorNip); }

// ==========================================
// FUNGSI ADMIN / SUPER ADMIN
// ==========================================
function getAllDataKomper(actorNip) {
  if (!_isAdminOrSuperAdmin(actorNip)) {
    return [];
  }
  const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headersRaw = data[0];
  const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
  const result = [];
  
  const passIndex = headers.indexOf("PASSWORD");

  for (let i = 1; i < data.length; i++) {
    let rowObj = {};
    
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      
      if (val instanceof Date) {
        val = val.toString();
      }
      
      rowObj[headers[j]] = val !== undefined ? val : "";
    }

    // KEAMANAN (FASE 3): Masking Password di API
    if (passIndex !== -1) {
      rowObj["PASSWORD"] = "******";
    }
    
    // OVERRIDE REAL-TIME STATUS KOMPETENSI UNTUK ADMIN
    rowObj["STATUS STR"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU STR AKTIF"], rowObj["STATUS STR"]);
    rowObj["STATUS SIPP SIPB"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU SIPP SIPB AKTIF"], rowObj["STATUS SIPP SIPB"]);
    rowObj["STATUS BTCLS"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU BTCLS"], rowObj["STATUS BTCLS"]);
    rowObj["STATUS PELATIHAN KAMAR OPERASI"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU PELATIHAN KAMAR OPERASI"], rowObj["STATUS PELATIHAN KAMAR OPERASI"]);
    rowObj["STATUS PELATIHAN INTENSIF"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU PELATIHAN INTENSIF"], rowObj["STATUS PELATIHAN INTENSIF"]);
    rowObj["STATUS PELATIHAN HD"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU PELATIHAN HD"], rowObj["STATUS PELATIHAN HD"]);
    rowObj["STATUS PELATIHAN MU"] = calculateRealtimeStatusServer(rowObj["MASA BERLAKU PELATIHAN MIDWIFE UPDATE"], rowObj["STATUS PELATIHAN MU"]);

    let calculatedKredensial = calculateKredensialStatusServer(rowObj["TANGGAL KREDENSIAL"]);
    if (calculatedKredensial !== "-") {
      rowObj["STATUS KREDENSIAL"] = calculatedKredensial;
    }

    result.push(rowObj);
  }
  return result;
}

function updateVerifikasiStatus(nip, status, actorNip) {
  if (!_isAdminOrSuperAdmin(actorNip)) {
    return { success: false, message: "Akses ditolak: Hanya Admin/Super Admin yang diizinkan mengubah status verifikasi." };
  }

  const ALLOWED_STATUS = ["Terverifikasi", "Belum Ter Verifikasi", "Ditolak"];
  if (ALLOWED_STATUS.indexOf(status) === -1) {
    return { success: false, message: "Status verifikasi tidak valid." };
  }

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      return { success: false, message: "Server sedang sibuk. Silakan coba sesaat lagi." };
    }

    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
    const data = sheet.getDataRange().getValues();
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    
    const nipIndex = headers.indexOf("NIP");
    const verifIndex = headers.indexOf("VERIFIKASI");
    
    if (nipIndex === -1 || verifIndex === -1) {
      return { success: false, message: "Kolom NIP atau VERIFIKASI tidak ditemukan di Sheet DATA KOMPER." };
    }
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIndex]).trim() === String(nip).trim()) {
        sheet.getRange(i + 1, verifIndex + 1).setValue(status);
        return { success: true, message: "Status Verifikasi berhasil diperbarui." };
      }
    }
    return { success: false, message: "Pegawai dengan NIP tersebut tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(lErr) {}
  }
}

function tambahDataPegawai(d, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
    const data = sheet.getDataRange().getValues();
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    
    const nipIndex = headers.indexOf("NIP");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIndex]).trim() === String(d.nip).trim()) {
        return { success: false, message: "NIP sudah terdaftar dalam sistem." };
      }
    }

    let newRow = new Array(headers.length).fill('');
    newRow[headers.indexOf("NIP")] = d.nip;
    newRow[headers.indexOf("NAMA")] = d.nama;
    newRow[headers.indexOf("PASSWORD")] = d.password;
    newRow[headers.indexOf("ROLE")] = d.role;
    newRow[headers.indexOf("VERIFIKASI")] = "Belum Ter Verifikasi";

    sheet.appendRow(newRow);
    return { success: true, message: "Data berhasil ditambahkan." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function mutasiPegawai(nip, actorNip) {
  try {
    if (!actorNip || !_isAdminOrSuperAdmin(actorNip)) {
      return { success: false, message: "Akses ditolak: Membutuhkan role Admin/Super Admin." };
    }

    var ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    var wsKomper = ss.getSheetByName("DATA KOMPER");
    var wsMutasi = ss.getSheetByName("MUTASI"); 
    
    if (!wsMutasi) {
      return { success: false, message: "Sheet MUTASI belum dibuat di Spreadsheet." };
    }
    
    var data = wsKomper.getDataRange().getValues();
    var headers = data[0];
    var nipIndex = headers.indexOf("NIP");
    
    if (nipIndex === -1) return { success: false, message: "Kolom NIP tidak ditemukan." };
    
    var rowIndex = -1;
    var rowData = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][nipIndex].toString() === nip.toString()) {
        rowData = data[i];
        rowIndex = i + 1; // +1 karena index array mulai dari 0, baris sheet mulai dari 1
        break;
      }
    }
    
    if (rowIndex > -1) {
      wsMutasi.appendRow(rowData); // Kirim seluruh baris ke sheet MUTASI
      wsKomper.deleteRow(rowIndex); // Hapus baris dari sheet DATA KOMPER
      return { success: true, message: "Data berhasil dimutasi." };
    } else {
      return { success: false, message: "Data pegawai tidak ditemukan." };
    }
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// FUNGSI PEMBERSIH FILE (AUTO-DESTRUCT)
// ==========================================
function triggerCleanup() {
  try {
    // 1. Cari file ekspor yang belum terhapus
    const oldFiles = DriveApp.searchFiles("title contains 'Kredensial_Export_' and trashed = false");
    const now = new Date().getTime();
    while (oldFiles.hasNext()) {
      let f = oldFiles.next();
      // 2. Hapus file jika umurnya lebih dari 2 menit (120000 ms)
      if (now - f.getDateCreated().getTime() > 120000) {
        f.setTrashed(true);
      }
    }
    
    // 3. Bersihkan trigger diri sendiri agar kuota (max 20 trigger) tidak penuh
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'triggerCleanup') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {}
}

// ==========================================
// FUNGSI EXPORT KREDENSIAL (WORD TEMPLATE)
// ==========================================
function exportKredensial(nip, actorNip) {
  try {
    const effectiveActor = actorNip || nip;
    if (!_canEditNip(nip, effectiveActor)) {
      return { success: false, message: "Akses ditolak: Anda tidak memiliki wewenang untuk mengekspor dokumen pegawai lain." };
    }

    // --- 1. CLEANUP FILE SEMENTARA SEBELUMNYA SECARA AGRESIF ---
    try {
      const oldFiles = DriveApp.searchFiles("title contains 'Kredensial_Export_' and trashed = false");
      const now = new Date().getTime();
      while (oldFiles.hasNext()) {
        let f = oldFiles.next();
        if (now - f.getDateCreated().getTime() > 60000) { // 1 menit
          f.setTrashed(true);
        }
      }
    } catch(err) {}

    // --- 2. AMBIL DATA PEGAWAI DARI SPREADSHEET ---
    const ssId = "1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA";
    const sheet = SpreadsheetApp.openById(ssId).getSheetByName("DATA KOMPER");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h).trim().toUpperCase(); });
    
    const nipIdx = headers.indexOf("NIP");
    let userData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIdx]).trim() === String(nip).trim()) {
        userData = data[i];
        break;
      }
    }
    
    if (!userData) return { success: false, message: "Data pegawai tidak ditemukan." };

    // --- 3. DUPLIKASI TEMPLATE & INISIALISASI DOKUMEN ---
    const templateId = "11ZBLjpRmeZYBTrR9aZVDZL4yX1clju6A5eWhfw7B07s";
    const namaPegawai = String(userData[headers.indexOf("NAMA")] || nip);
    
    let tempDoc;
    try {
        tempDoc = DriveApp.getFileById(templateId).makeCopy("Kredensial_Export_" + namaPegawai);
    } catch(errTemplate) {
        return { 
            success: false, 
            message: "Gagal mengakses Template. Solusi: 1) Pastikan ID Template benar. 2) Buka Google Drive, Share file template menjadi 'Siapa saja yang memiliki link (Viewer)'. 3) Pastikan Web App di-deploy dengan 'Execute as: Me'. Detail Error: " + errTemplate.message 
        };
    }
    const docId = tempDoc.getId();
    
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();
    
    const formatDate = function(val) {
       if (val instanceof Date) {
           const y = val.getFullYear();
           const m = String(val.getMonth() + 1).padStart(2, '0');
           const d = String(val.getDate()).padStart(2, '0');
           return d + '-' + m + '-' + y;
       }
       return val ? String(val) : "";
    };

    const getVal = function(colName) {
       const target = String(colName).trim().toUpperCase();
       const idx = headers.indexOf(target);
       return (idx > -1 && userData[idx] !== undefined && userData[idx] !== null) ? userData[idx] : "";
    };

    // Ekstraksi 4 digit Tahun Lulus
    let tahunLulus = getVal("TAHUN LULUS PENDIDIKAN TERAKHIR");
    if (tahunLulus) {
        let match = String(tahunLulus).match(/\b(19|20)\d{2}\b/);
        if (match) tahunLulus = match[0];
    }

    // --- 4. PEMETAAN TEXT ---
    const mapData = {
        "NAMA": getVal("NAMA"),
        "TEMPAT_LAHIR": getVal("TEMPAT LAHIR"),
        "TANGGAL_LAHIR": formatDate(getVal("TANGGAL LAHIR")),
        "JENIS_KELAMIN": getVal("JENIS KELAMIN"),
        "NOMOR_HANDPHONE": getVal("NOMOR HANDPHONE"),
        "ALAMAT_EMAIL": getVal("ALAMAT EMAIL"),
        "INSTITUSI_PENDIDIKAN_TERAKHIR": getVal("INSTITUSI PENDIDIKAN TERAKHIR"),
        "JURUSAN": getVal("JURUSAN"),
        "JENJANG_PENDIDIKAN_TERAKHIR": getVal("JENJANG PENDIDIKAN TERAKHIR"),
        "TAHUN_LULUS_PENDIDIKAN_TERAKHIR": tahunLulus,
        "JABATAN": getVal("JABATAN"),
        "NIP": getVal("NIP")
    };

    // Ekstraksi ID Foto secara aman dari URL Google Drive
    let fotoUrl = getVal("FOTO");
    let fotoId = null;
    if (fotoUrl && fotoUrl !== "-" && fotoUrl !== "") {
        let strUrl = String(fotoUrl);
        let match = strUrl.match(/id=([-\w]{15,})/); // Ubah batasan length regex untuk menoleransi berbagai ID Google Drive
        if (!match) match = strUrl.match(/\/d\/([-\w]{15,})/);
        if (match && match[1]) {
            fotoId = String(match[1]).replace(/[^a-zA-Z0-9_-]/g, ""); // Anti hidden-characters pada ID
        }
    }

    // Ambil foto blob SATU KALI saja untuk mencegah error pemanggilan berkali-kali
    let fotoBlob = null;
    if (fotoId && fotoId !== "") {
        try {
            fotoBlob = DriveApp.getFileById(fotoId).getBlob();
        } catch (e) {
            fotoBlob = null;
        }
    }

    const elementsToReplace = [];
    if (body) elementsToReplace.push(body);
    if (header) elementsToReplace.push(header);
    if (footer) elementsToReplace.push(footer);

    elementsToReplace.forEach(function(element) {
        if (element) {
            // A. Mengganti semua variabel Teks dengan Regex Anti-Error
            for (let key in mapData) {
                let val = mapData[key];
                if (val === undefined || val === null) val = "";
                
                // Menangani format <<TAG>>, «TAG», << TAG >> yang sering rusak oleh Google Docs
                element.replaceText("<<" + key + ">>", String(val)); 
                element.replaceText("«" + key + "»", String(val)); 
                element.replaceText("<<\\s*" + key + "\\s*>>", String(val)); 
                element.replaceText("«\\s*" + key + "\\s*»", String(val)); 
            }

            // B. Proses Penempelan Gambar (FOTO) - Mengubah string <<FOTO>> menjadi gambar fisik
            let fotoTags = ["<<FOTO>>", "«FOTO»", "<< FOTO >>", "« FOTO »", "<<\\s*FOTO\\s*>>"];
            if (fotoBlob) {
                fotoTags.forEach(function(fTag) {
                    let found = element.findText(fTag);
                    while (found) {
                        let textEl = found.getElement().asText();
                        let p = textEl.getParent();
                        
                        // Sisipkan gambar asli seukuran pasfoto 3x4 (85x113 pt) agar rapih di tabel
                        if (p.getType() === DocumentApp.ElementType.PARAGRAPH) {
                            let img = p.asParagraph().insertInlineImage(p.getChildIndex(textEl), fotoBlob);
                            img.setWidth(85); 
                            img.setHeight(113);
                        } else if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
                            let img = p.asListItem().insertInlineImage(p.getChildIndex(textEl), fotoBlob);
                            img.setWidth(85); 
                            img.setHeight(113);
                        }
                        
                        // Hapus label text <<FOTO>> 
                        textEl.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
                        
                        // Bersihkan spasi sisa agar posisi foto tidak tergeser / berantakan
                        if (textEl.getText().trim() === "") {
                            textEl.setText(""); 
                        }
                        
                        found = element.findText(fTag); 
                    }
                });
            } else {
                // Hapus label jika tidak ada foto / gagal diakses
                fotoTags.forEach(tag => element.replaceText(tag, ""));
            }
        }
    });
    
    doc.saveAndClose();
    
    // --- 5. GENERATE URL DOWNLOAD & AKSES ---
    tempDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const exportUrl = "https://docs.google.com/document/d/" + docId + "/export?format=docx";
    
    // --- 6. JADWALKAN PENGHAPUSAN OTOMATIS (AUTO-DESTRUCT) ---
    // Memicu trigger untuk menghapus file ini dalam waktu 3 Menit
    try {
      ScriptApp.newTrigger("triggerCleanup")
               .timeBased()
               .after(3 * 60 * 1000) // 3 Menit
               .create();
    } catch (e) {
      // Jika gagal membuat trigger (batas harian tercapai/otorisasi), sistem akan
      // tetap menghapusnya pada eksekusi export berikutnya (langkah ke-1).
    }

    return { success: true, url: exportUrl };
  } catch (e) {
    return { success: false, message: "Sistem error saat memproses dokumen: " + e.message };
  }
}// ==========================================
// PENGATURAN USERS & SLIDESHOW (SETTING MENU)
// ==========================================

function getUsersList(actorNip) {
  if (!_isSuperAdmin(actorNip)) return [];
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    
    // Baca OTORISASI
    let mapAkses = {};
    try {
      const sheetOtorisasi = ss.getSheetByName("OTORISASI");
      if (sheetOtorisasi) {
        const dataOto = sheetOtorisasi.getDataRange().getValues();
          for (let i = 1; i < dataOto.length; i++) {
            let nipOto = String(dataOto[i][0]).trim();
            let menuStr = String(dataOto[i][1]).trim();
            if (nipOto) mapAkses[nipOto] = menuStr ? menuStr.split(",").map(function(s){return s.trim()}).filter(Boolean) : [];
          }
      }
    } catch(e) {}
    
    const sheet = ss.getSheetByName("DATA KOMPER");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    const result = [];
    
    const idxNip = headers.indexOf("NIP");
    const idxNama = headers.indexOf("NAMA");
    const idxEmail = headers.indexOf("ALAMAT EMAIL");
    const idxHp = headers.indexOf("NOMOR HANDPHONE");
    const idxRole = headers.indexOf("ROLE");
    const idxVerif = headers.indexOf("VERIFIKASI");

    for (let i = 1; i < data.length; i++) {
      let nip = idxNip !== -1 ? String(data[i][idxNip]) : "";
      if (!nip) continue;
      
      result.push({
        NIP: nip,
        NAMA: idxNama !== -1 ? data[i][idxNama] : "",
        EMAIL: idxEmail !== -1 ? data[i][idxEmail] : "",
        HP: idxHp !== -1 ? data[i][idxHp] : "",
        ROLE: idxRole !== -1 ? data[i][idxRole] : "",
        VERIFIKASI: idxVerif !== -1 ? data[i][idxVerif] : "",
        MENU_AKSES: mapAkses[nip] !== undefined ? mapAkses[nip] : null
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

function updateUserAccount(nipToEdit, fieldToEdit, newValue, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    const sheet = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA").getSheetByName("DATA KOMPER");
    const data = sheet.getDataRange().getValues();
    const headersRaw = data[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    const nipIndex = headers.indexOf("NIP");
    const targetColIndex = headers.indexOf(String(fieldToEdit).toUpperCase());
    
    if (targetColIndex === -1) return { success: false, message: "Kolom " + fieldToEdit + " tidak ditemukan." };
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nipIndex]).trim() === String(nipToEdit).trim()) {
        
        // Safety guard: Jika merubah role
        if (String(fieldToEdit).toUpperCase() === "ROLE") {
          const roleIndex = headers.indexOf("ROLE");
          const oldRole = String(data[i][roleIndex]).trim();
          
          if (oldRole === "Super Admin" && newValue !== "Super Admin") {
            if (String(nipToEdit).trim() === String(actorNip).trim()) {
              return { success: false, message: "Anda tidak dapat menurunkan role Super Admin dari akun Anda sendiri." };
            }
          }
        }
        
        sheet.getRange(i + 1, targetColIndex + 1).setValue(newValue);
        return { success: true, message: fieldToEdit + " berhasil diperbarui." };
      }
    }
    return { success: false, message: "User tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _initSlideshowSheet(ss) {
  let sheet = ss.getSheetByName("SLIDESHOW");
  if (!sheet) {
    sheet = ss.insertSheet("SLIDESHOW");
    sheet.appendRow(["ID", "URL", "TANGGAL_DITAMBAHKAN"]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e0f2f1");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 600);
    sheet.setColumnWidth(3, 200);
  }
  return sheet;
}

function getSlideshowList() {
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initSlideshowSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    // Ambil URL dari folder drive (sebagai read-only list)
    const folderImages = [];
    try {
      const folderId = '1J3R4trtlJ9QaabMWYLC-c12ieYVEeUTX';
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getMimeType().indexOf('image/') !== -1) {
          folderImages.push({ id: file.getId(), url: file.getUrl(), source: 'drive' });
        }
      }
    } catch(e) {}

    const sheetImages = [];
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] && data[i][1]) {
        sheetImages.push({
          id: data[i][0],
          url: data[i][1],
          tanggal: data[i][2],
          source: 'sheet'
        });
      }
    }
    
    return folderImages.concat(sheetImages);
  } catch(e) {
    return [];
  }
}

function addSlideshowUrl(url, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    if(!url || !url.startsWith("http")) return { success: false, message: "URL tidak valid." };
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initSlideshowSheet(ss);
    const id = "IMG_" + new Date().getTime();
    const dateStr = new Date().toString();
    sheet.appendRow([id, url, dateStr]);
    return { success: true, message: "Gambar berhasil ditambahkan." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteSlideshowUrl(id, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    if(!id) return { success: false, message: "ID tidak valid." };
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initSlideshowSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Gambar berhasil dihapus." };
      }
    }
    return { success: false, message: "ID tidak ditemukan di database." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
function _initOtorisasiSheet(ss) {
  let sheet = ss.getSheetByName("OTORISASI");
  if (!sheet) {
    sheet = ss.insertSheet("OTORISASI");
    sheet.appendRow(["NIP", "MENU_AKSES", "DIPERBARUI"]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e0f2f1");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 400);
    sheet.setColumnWidth(3, 200);
  }
  return sheet;
}

function getUserAccess(nip) {
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initOtorisasiSheet(ss);
    const data = sheet.getDataRange().getValues();
    
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(nip).trim()) {
          const menuStr = String(data[i][1]).trim();
          return { success: true, menus: menuStr ? menuStr.split(",").map(function(s){return s.trim()}).filter(Boolean) : [] };
        }
      }
    // Jika belum pernah diatur, return null (berarti akses penuh)
    return { success: true, menus: null };
  } catch(e) {
    return { success: false, message: e.message, menus: null };
  }
}

function updateUserAccess(nipToEdit, menusArray, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initOtorisasiSheet(ss);
    const data = sheet.getDataRange().getValues();
    const dateStr = new Date().toString();
    const menuStr = (menusArray || []).join(",");
    
    // Cari baris NIP
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nipToEdit).trim()) {
        sheet.getRange(i + 1, 2).setValue(menuStr);
        sheet.getRange(i + 1, 3).setValue(dateStr);
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow([nipToEdit, menuStr, dateStr]);
    }
    return { success: true, message: "Otorisasi menu berhasil diperbarui." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateUserRoleAndAccess(nipToEdit, newRole, menusArray, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheetKomper = ss.getSheetByName("DATA KOMPER");
    const sheetOto = _initOtorisasiSheet(ss);
    
    // Update ROLE in DATA KOMPER
    const dataK = sheetKomper.getDataRange().getValues();
    const headersRaw = dataK[0];
    const headers = headersRaw.map(function(h) { return String(h).trim().toUpperCase(); });
    const nipIndex = headers.indexOf("NIP");
    const roleIndex = headers.indexOf("ROLE");
    
    if (roleIndex === -1) return { success: false, message: "Kolom ROLE tidak ditemukan." };
    
    let userFound = false;
    for (let i = 1; i < dataK.length; i++) {
      if (String(dataK[i][nipIndex]).trim() === String(nipToEdit).trim()) {
        const oldRole = String(dataK[i][roleIndex]).trim();
        if (oldRole === "Super Admin" && newRole !== "Super Admin") {
          if (String(nipToEdit).trim() === String(actorNip).trim()) {
            return { success: false, message: "Anda tidak dapat menurunkan role Super Admin dari akun Anda sendiri." };
          }
        }
        sheetKomper.getRange(i + 1, roleIndex + 1).setValue(newRole);
        userFound = true;
        break;
      }
    }
    
    if (!userFound) return { success: false, message: "User tidak ditemukan di DATA KOMPER." };

    // Update OTORISASI
    const dataOto = sheetOto.getDataRange().getValues();
    const dateStr = new Date().toString();
    const menuStr = (menusArray || []).join(",");
    let otoFound = false;
    
    for (let i = 1; i < dataOto.length; i++) {
      if (String(dataOto[i][0]).trim() === String(nipToEdit).trim()) {
        sheetOto.getRange(i + 1, 2).setValue(menuStr);
        sheetOto.getRange(i + 1, 3).setValue(dateStr);
        otoFound = true;
        break;
      }
    }
    if (!otoFound) {
      sheetOto.appendRow([nipToEdit, menuStr, dateStr]);
    }
    
    return { success: true, message: "Role dan Otorisasi berhasil diperbarui." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function uploadSlideshowImage(base64Data, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  
  try {
    let type = 'application/octet-stream';
    let base64String = base64Data;
    if (base64Data.indexOf(',') !== -1) {
      const splitBase = base64Data.split(',');
      type = splitBase[0].split(';')[0].replace('data:', '');
      base64String = splitBase[1];
    }
    
    const byteCharacters = Utilities.base64Decode(base64String);
    const blob = Utilities.newBlob(byteCharacters, type, "temp");
    
    // Validasi ukuran maksimal 3 MB
    if (blob.getBytes().length > 3 * 1024 * 1024) {
      return { success: false, message: "Ukuran file gambar melebihi batas maksimal 3 MB." };
    }
  } catch (e) {
    // Abaikan jika tipe bukan image atau gagal decode, biarkan genericUpload yang menangani errornya
  }

  const FOLDER_ID = '1J3R4trtlJ9QaabMWYLC-c12ieYVEeUTX';
  return genericUpload(base64Data, actorNip, FOLDER_ID, 'Slideshow', null);
}

function deleteSlideshowDriveFile(id, actorNip) {
  if (!_isSuperAdmin(actorNip)) return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  try {
    if(!id) return { success: false, message: "ID tidak valid." };
    const file = DriveApp.getFileById(id);
    file.setTrashed(true);
    return { success: true, message: "File gambar berhasil dihapus dari Google Drive." };
  } catch(e) {
    return { success: false, message: "Gagal menghapus file dari Drive: " + e.message };
  }
}

// ==========================================
// PENGATURAN SISTEM (RUNNING TEXT)
// ==========================================
function _initPengaturanSheet(ss) {
  let sheet = ss.getSheetByName("PENGATURAN");
  if (!sheet) {
    sheet = ss.insertSheet("PENGATURAN");
    sheet.appendRow(["KEY", "NILAI", "DIPERBARUI"]);
    sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#e0f2f1");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 600);
    sheet.setColumnWidth(3, 200);
    sheet.appendRow(["RUNNING_TEXT", "Selamat Datang di Sistem Informasi Komite Keperawatan RSUD Subang. Pastikan kelengkapan data kepegawaian dan kredensial Anda selalu terbarui.", new Date().toString()]);
  }
  return sheet;
}

function getRunningText() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("RUNNING_TEXT");
  if (cached !== null) {
    return cached;
  }

  try {
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initPengaturanSheet(ss);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "RUNNING_TEXT") {
        const val = String(data[i][1] || "").trim();
        try { cache.put("RUNNING_TEXT", val, 21600); } catch(cErr) {}
        return val;
      }
    }
    return "";
  } catch(e) {
    return "";
  }
}

function saveRunningText(teks, actorNip) {
  if (!_isSuperAdmin(actorNip)) {
    return { success: false, message: "Akses ditolak: Hanya Super Admin yang diizinkan." };
  }
  try {
    const cleanTeks = String(teks || "").trim().slice(0, 500);
    const ss = SpreadsheetApp.openById("1fuuoaGFVBFqK3gOKHQFsB5Kr5dqUJSehlCdtnvyrOhA");
    const sheet = _initPengaturanSheet(ss);
    const data = sheet.getDataRange().getValues();
    const dateStr = new Date().toString();
    
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "RUNNING_TEXT") {
        sheet.getRange(i + 1, 2).setValue(cleanTeks);
        sheet.getRange(i + 1, 3).setValue(dateStr);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(["RUNNING_TEXT", cleanTeks, dateStr]);
    }
    try {
      CacheService.getScriptCache().put("RUNNING_TEXT", cleanTeks, 21600);
    } catch(cErr) {}
    return { success: true, message: "Running Text berhasil disimpan." };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan Running Text: " + e.message };
  }
}

