// ============================================================
//  User.gs — semua fungsi yang dipanggil guru/tendik untuk
//  dirinya sendiri (laporan, izin, kegiatan, data, profil)
// ============================================================

// ─── LAPORAN MENGAJAR ─────────────────────────────────────
// Aturan:
//  1. Kelas & mata pelajaran dipilih dari Master Data (kosong = boleh ketik bebas).
//  2. Jam mulai & selesai wajib diisi (dipakai untuk cek tabrakan).
//  3. Milik guru yang SAMA, di tanggal yang sama, jam tidak boleh beririsan dengan laporan
//     yang sudah ada — apa pun kelas/mapelnya. (Guru LAIN tidak terpengaruh.)
//     Kombinasi guru+kelas+mapel+jam yang sama otomatis tertolak karena pasti beririsan.
//  4. Pengecekan + penyimpanan berjalan di dalam LockService, jadi klik ganda / aplikasi
//     lambat / dua request bersamaan tidak bisa menyimpan laporan dua kali.
// Catatan: kolom Materi, JumlahSiswa, dan Keterangan sudah tidak diisi dari form (disimpan
// kosong/0) — kolomnya tetap ada di sheet supaya data lama tidak berubah posisi/format.
const REGEX_JAM = /^([01]?\d|2[0-3]):[0-5]\d$/;
const REGEX_TANGGAL = /^\d{4}-\d{2}-\d{2}$/;

/** yyyy-MM-dd dan benar-benar ada di kalender (2026-02-31 ditolak, bukan dibulatkan ke 3 Maret). */
function tanggalValid_(t) {
  if (!REGEX_TANGGAL.test(t)) return false;
  const [y, m, d] = t.split('-').map(Number);
  const x = new Date(Date.UTC(y, m - 1, d));
  return x.getUTCFullYear() === y && x.getUTCMonth() === m - 1 && x.getUTCDate() === d;
}

/** Laporan milik satu guru di satu tanggal (yyyy-MM-dd, WIB), jam sudah dinormalisasi. */
function laporanUserTanggal_(userID, tanggal) {
  const tz = 'Asia/Jakarta';
  return getSheetData(SHEET_NAMES.LAPORAN)
    .filter(l => l.UserID === userID && l.Tanggal &&
      Utilities.formatDate(new Date(l.Tanggal), tz, 'yyyy-MM-dd') === tanggal)
    .map(l => {
      const m = jamKeMenit(l.JamMulai), s = jamKeMenit(l.JamSelesai);
      const valid = l.JamMulai && l.JamSelesai && s > m;
      return {
        mulaiMenit: m, selesaiMenit: s, menit: valid ? s - m : 0,
        mulai: menitKeStr(m), selesai: menitKeStr(s),
        mapel: l.Mapel || '', kelas: l.Kelas || ''
      };
    })
    .sort((a, b) => a.mulaiMenit - b.mulaiMenit);
}

function submitLaporanMengajar(token, data) {
  const lock = LockService.getScriptLock();
  let terkunci = false;
  try {
    const session = requireUser(token);
    const userID = session.userID; // [KEAMANAN] Pakai userID dari session, bukan dari client

    // Setelah absen pulang, hari itu dianggap selesai — berlaku untuk semua role.
    if (cekStatusAbsenHariIni(token).pulang) {
      return { success: false, message: 'Anda sudah absen pulang hari ini. Laporan Mengajar tidak bisa diisi lagi untuk hari ini.' };
    }
    const kunciWaktu = statusKunciWaktu_();
    if (kunciWaktu.lewat) return { success: false, message: `Waktu pengisian Laporan Mengajar untuk hari ini sudah ditutup sejak pukul ${kunciWaktu.jamKunci}. Hubungi admin bila perlu mengisi data untuk hari ini.` };

    // Sanitasi input
    const tanggal    = sanitize(data.tanggal);
    const jamMulai   = sanitize(data.jamMulai);
    const jamSelesai = sanitize(data.jamSelesai);

    // Kelas & mapel: cocokkan dengan Master Data (jangan disanitasi dulu, agar "Qur'an Hadis" tetap cocok)
    const kelasM = cocokkanMaster_('Kelas', data.kelas, 'Kelas');
    if (!kelasM.ok) return { success: false, message: kelasM.pesan };
    const mapelM = cocokkanMaster_('Mapel', data.mapel, 'Mata pelajaran');
    if (!mapelM.ok) return { success: false, message: mapelM.pesan };
    const kelas = kelasM.nilai, mapel = mapelM.nilai;

    if (!tanggalValid_(tanggal)) return { success: false, message: 'Tanggal tidak valid' };

    // Jam wajib & valid
    if (!REGEX_JAM.test(jamMulai) || !REGEX_JAM.test(jamSelesai)) {
      return { success: false, message: 'Jam mulai dan jam selesai wajib diisi.' };
    }
    const mulai = jamKeMenit(jamMulai), selesai = jamKeMenit(jamSelesai);
    if (selesai <= mulai) return { success: false, message: 'Jam selesai harus lebih dari jam mulai' };
    const durasi = selesai - mulai;
    if (durasi > 480) return { success: false, message: 'Durasi mengajar tidak boleh lebih dari 8 jam' };

    try { lock.waitLock(15000); }
    catch (err) { return { success: false, message: 'Sistem sedang sibuk. Tunggu sebentar, cek dulu riwayat laporan, baru kirim ulang bila belum tersimpan.' }; }
    terkunci = true;

    // Cek tabrakan dengan laporan guru ini di tanggal yang sama
    const bentrok = laporanUserTanggal_(userID, tanggal)
      .find(l => l.menit > 0 && mulai < l.selesaiMenit && selesai > l.mulaiMenit);
    if (bentrok) {
      return { success: false, message: `Jam ${menitKeStr(mulai)}-${menitKeStr(selesai)} bertabrakan dengan laporan yang sudah Anda isi pada tanggal ini: ${bentrok.mulai}-${bentrok.selesai} (${bentrok.mapel || '-'}, ${bentrok.kelas || '-'}).\nJam mengajar yang sama tidak bisa diisi dua kali.` };
    }

    const users = getSheetData(SHEET_NAMES.USER);
    const user = users.find(u => u.ID === userID);
    const nama = user ? user.Nama : 'Unknown';

    const now = new Date();
    const id = generateID('LAP');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAMES.LAPORAN);
    const baris = sh.getLastRow() + 1;
    sh.getRange(baris, 8, 1, 2).setNumberFormat('@'); // jam disimpan sebagai TEKS, bukan diubah Sheets jadi waktu
    sh.getRange(baris, 1, 1, 13).setValues([[id, userID, nama, new Date(tanggal), kelas, mapel, '', menitKeStr(mulai), menitKeStr(selesai), 0, '', data.absID || '', now]]);
    addLog(userID, nama, 'LAPORAN_MENGAJAR', `Kelas: ${kelas} | Mapel: ${mapel} | Durasi: ${durasi} menit`);
    return { success: true, message: 'Laporan mengajar berhasil disimpan!' };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    if (terkunci) lock.releaseLock();
  }
}

// ─── PENGAJUAN IZIN ───────────────────────────────────────

// ─── PENGAJUAN IZIN ───────────────────────────────────────
function submitIzin(token, data) {
  try {
    const session = requireUser(token);
    const userID = session.userID;

    const jenisIzin   = sanitize(data.jenisIzin);
    const keterangan  = sanitize(data.keterangan);
    const tglMulai    = sanitize(data.tglMulai);
    const tglSelesai  = sanitize(data.tglSelesai);

    if (!keterangan) return { success: false, message: 'Keterangan izin wajib diisi' };

    const users = getSheetData(SHEET_NAMES.USER);
    const user = users.find(u => u.ID === userID);
    const nama = user ? user.Nama : 'Unknown';

    const id = generateID('IZN');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAMES.IZIN);
    sh.appendRow([id, userID, nama, new Date(), new Date(tglMulai), new Date(tglSelesai), jenisIzin, keterangan, '', 'MENUNGGU', '', '', '']);
    addLog(userID, nama, 'PENGAJUAN_IZIN', `Jenis: ${jenisIzin} | ${tglMulai} s/d ${tglSelesai}`);
    return { success: true, message: 'Pengajuan izin berhasil dikirim, menunggu persetujuan admin.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── PENGAJUAN KEGIATAN ───────────────────────────────────
// Aturan:
//  1. Nama kegiatan dipilih dari Master Data (kosong = boleh ketik bebas).
//  2. Langsung berstatus DISETUJUI — tidak ada lagi persetujuan admin.
//  3. Guru yang SAMA tidak bisa mengajukan kegiatan yang SAMA di tanggal yang SAMA.
//     (Guru lain, kegiatan lain, atau tanggal lain tetap boleh.)
//  4. Pengecekan + penyimpanan berjalan di dalam LockService, jadi klik ganda / aplikasi
//     lambat tidak bisa membuat pengajuan dobel.
function kegiatanUserTanggal_(userID, tanggal) {
  const tz = 'Asia/Jakarta';
  return getSheetData(SHEET_NAMES.KEGIATAN)
    .filter(k => k.UserID === userID && k.TanggalKegiatan &&
      Utilities.formatDate(new Date(k.TanggalKegiatan), tz, 'yyyy-MM-dd') === tanggal)
    .map(k => String(k.NamaKegiatan || '').trim().toLowerCase());
}

function submitKegiatan(token, data) {
  const lock = LockService.getScriptLock();
  let terkunci = false;
  try {
    const session = requireUser(token);
    const userID = session.userID;

    // Setelah absen pulang, hari itu dianggap selesai — berlaku untuk semua role.
    if (cekStatusAbsenHariIni(token).pulang) {
      return { success: false, message: 'Anda sudah absen pulang hari ini. Kegiatan tidak bisa diisi lagi untuk hari ini.' };
    }
    const kunciWaktu = statusKunciWaktu_();
    if (kunciWaktu.lewat) return { success: false, message: `Waktu pengisian Kegiatan untuk hari ini sudah ditutup sejak pukul ${kunciWaktu.jamKunci}. Hubungi admin bila perlu mengisi data untuk hari ini.` };

    const tanggal = sanitize(data.tanggal);
    const lokasi  = sanitize(data.lokasi);
    const deskripsi = sanitize(data.deskripsi);

    const kegM = cocokkanMaster_('Kegiatan', data.namaKegiatan, 'Kegiatan');
    if (!kegM.ok) return { success: false, message: kegM.pesan };
    const namaKegiatan = kegM.nilai;

    if (!tanggalValid_(tanggal)) return { success: false, message: 'Tanggal tidak valid' };

    try { lock.waitLock(15000); }
    catch (err) { return { success: false, message: 'Sistem sedang sibuk. Tunggu sebentar, cek dulu riwayat kegiatan, baru kirim ulang bila belum tersimpan.' }; }
    terkunci = true;

    if (kegiatanUserTanggal_(userID, tanggal).includes(namaKegiatan.toLowerCase())) {
      return { success: false, message: `Anda sudah mengajukan kegiatan "${namaKegiatan}" pada tanggal ini. Kegiatan yang sama tidak bisa diajukan dua kali di hari yang sama.` };
    }

    const users = getSheetData(SHEET_NAMES.USER);
    const user = users.find(u => u.ID === userID);
    const nama = user ? user.Nama : 'Unknown';

    const id = generateID('KEG');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAMES.KEGIATAN);
    sh.appendRow([id, userID, nama, new Date(), new Date(tanggal), '', namaKegiatan, lokasi, deskripsi, 'DISETUJUI', '', new Date(), '']);
    addLog(userID, nama, 'PENGAJUAN_KEGIATAN', `Kegiatan: ${namaKegiatan}`);
    return { success: true, message: 'Kegiatan berhasil dicatat.' };
  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    if (terkunci) lock.releaseLock();
  }
}

// ─── GURU: PILIHAN DROPDOWN FORM KEGIATAN ─────────────────
function getOpsiKegiatan(token) {
  try {
    requireUser(token);
    return { success: true, kegiatan: getMasterList_('Kegiatan').map(x => x.nama) };
  } catch (e) { return { success: false, message: e.message }; }
}

// ─── STATUS LAPORAN HARI INI (dipakai untuk mengunci tombol Absen Pulang) ──
/**
 * Cek apakah user sudah punya minimal satu laporan mengajar hari ini.
 * Dipakai guru saja (role user_guru); role lain tidak wajib mengisi laporan.
 */
function getStatusLaporanHariIni(token) {
  try {
    const session = requireUser(token);
    const tz = 'Asia/Jakarta';
    const tanggal = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const jumlah = getSheetData(SHEET_NAMES.LAPORAN).filter(l =>
      l.UserID === session.userID && l.Tanggal &&
      Utilities.formatDate(new Date(l.Tanggal), tz, 'yyyy-MM-dd') === tanggal
    ).length;
    return { success: true, sudahLaporan: jumlah > 0, jumlah };
  } catch (e) { return { success: false, message: e.message }; }
}

// ─── STATUS AKSES HARIAN: LAPORAN MENGAJAR & KEGIATAN ─────
// Dipakai untuk mengunci SELURUH halaman Laporan Mengajar & Kegiatan:
//  - belum absen masuk           → terkunci (pesan: harus absen masuk dulu)
//  - sudah lewat jam_pulang+batas → terkunci (pesan: waktu pengisian ditutup)
// Berlaku untuk semua role (guru & tendik), tidak dibedakan seperti aturan Pulang.
function getStatusKunciHarian(token) {
  try {
    requireUser(token);
    const st = cekStatusAbsenHariIni(token);
    const kunci = statusKunciWaktu_();
    return { success: true, masuk: !!st.masuk, pulang: !!st.pulang, terkunciWaktu: kunci.lewat, jamKunci: kunci.jamKunci };
  } catch (e) { return { success: false, message: e.message }; }
}

// ─── RINGKASAN HARI INI (satu panggilan, dipakai halaman Beranda & Absensi) ──
// Menggabungkan cekStatusAbsenHariIni + getStatusLaporanHariIni + riwayat7Hari supaya cukup
// SATU round-trip ke server, bukan dua/tiga panggilan berurutan (lebih cepat).
function getRingkasanHariIni(token) {
  try {
    const session = requireUser(token);
    const userID = session.userID;
    const st = cekStatusAbsenHariIni(token);
    const lap = getStatusLaporanHariIni(token);

    const tz = 'Asia/Jakarta';
    // [FIX] Helper tanggal anti-crash: tanggal kosong/rusak dilewati, bukan melempar error
    // yang bikin seluruh getRingkasanHariIni gagal (dashboard user jadi kosong).
    const aman = v => { if (!v) return null; const d = new Date(v); return isNaN(d) ? null : d; };
    const jam = v => { const d = aman(v); return d ? Utilities.formatDate(d, tz, 'HH:mm') : '—'; };
    const tgl = v => { const d = aman(v); return d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; };

    // Ambil riwayat absensi 7 hari terakhir milik user
    const allAbs = getSheetData(SHEET_NAMES.ABSENSI);
    const riwayat7Hari = allAbs
      .filter(a => a.UserID === userID && aman(a.Tanggal))
      .sort((a, b) => new Date(b.Tanggal) - new Date(a.Tanggal))
      .slice(0, 7)
      .map(a => ({
        id: a.ID,
        tanggal: tgl(a.Tanggal),
        waktuMasuk: jam(a.WaktuMasuk),
        waktuPulang: jam(a.WaktuPulang),
        statusMasuk: a.StatusMasuk || '',
        statusPulang: a.StatusPulang || '',
        status: (a.StatusMasuk === 'TERLAMBAT' ? 'TERLAMBAT' : (a.StatusPulang === 'PULANG AWAL' ? 'PULANG AWAL' : (a.StatusMasuk || 'TEPAT WAKTU'))),
        lokasi: a.LokasiMasuk || a.Lokasi || '—'
      }));

    return {
      success: true,
      masuk: !!st.masuk, pulang: !!st.pulang,
      waktuMasuk: st.waktuMasuk || null, waktuPulang: st.waktuPulang || null,
      statusMasuk: st.statusMasuk || null, lokasi: st.lokasi || null,
      sudahLaporan: !!(lap && lap.success && lap.sudahLaporan),
      jumlahLaporan: (lap && lap.success) ? lap.jumlah : 0,
      riwayat7Hari
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ─── RINGKASAN BULANAN (statistik bulan berjalan milik user untuk Halaman Utama) ──
function getRingkasanBulananUser(token) {
  try {
    const session = requireUser(token);
    const userID = session.userID;
    const now = new Date();
    const bulan = now.getMonth() + 1;
    const tahun = now.getFullYear();
    const p = x => String(x).padStart(2, '0');
    const jumlahHari = new Date(Date.UTC(tahun, bulan, 0)).getUTCDate();
    const rentangAwal = `${tahun}-${p(bulan)}-01`;
    const rentangAkhir = `${tahun}-${p(bulan)}-${p(jumlahHari)}`;

    const tz = 'Asia/Jakarta';
    const tgl = v => v ? Utilities.formatDate(new Date(v), tz, 'yyyy-MM-dd') : '';
    const dalamBulan = s => s >= rentangAwal && s <= rentangAkhir;
    const ms = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));

    const liburs = (typeof HARI_LIBUR_MINGGUAN !== 'undefined') ? HARI_LIBUR_MINGGUAN : [0, 6];
    const bulans = (typeof NAMA_BULAN_REKAP !== 'undefined') ? NAMA_BULAN_REKAP : ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

    // Hari hadir & terlambat
    const absensi = getSheetData(SHEET_NAMES.ABSENSI);
    const absUser = absensi.filter(a => a.UserID === userID && a.WaktuMasuk && dalamBulan(tgl(a.Tanggal)));
    const hariHadir = {};
    absUser.forEach(a => {
      const t = tgl(a.Tanggal);
      const h = hariHadir[t] || (hariHadir[t] = { telat: false });
      if (a.StatusMasuk === 'TERLAMBAT') h.telat = true;
    });
    const hadir = Object.keys(hariHadir).length;
    const telat = Object.keys(hariHadir).filter(t => hariHadir[t].telat).length;

    // Izin yang disetujui bulan ini
    const izin = getSheetData(SHEET_NAMES.IZIN).filter(i => i.UserID === userID && i.Status === 'DISETUJUI');
    const terpakaiIzin = {};
    izin.forEach(i => {
      let mulai = tgl(i.TanggalMulai);
      let selesai = tgl(i.TanggalSelesai) || mulai;
      if (!mulai) return;
      if (mulai < rentangAwal) mulai = rentangAwal;
      if (selesai > rentangAkhir) selesai = rentangAkhir;
      for (let t = ms(mulai); t <= ms(selesai); t += 86400000) {
        const s = new Date(t).toISOString().slice(0, 10);
        if (liburs.includes(new Date(t).getUTCDay()) || hariHadir[s] || terpakaiIzin[s]) continue;
        terpakaiIzin[s] = true;
      }
    });
    const totalIzin = Object.keys(terpakaiIzin).length;

    // Hari kerja efektif yang sudah berlalu s/d hari ini
    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    let hkBerlalu = 0;
    for (let t = ms(rentangAwal); t <= Math.min(ms(rentangAkhir), ms(todayStr)); t += 86400000) {
      if (!liburs.includes(new Date(t).getUTCDay())) {
        hkBerlalu++;
      }
    }
    const alpha = Math.max(0, hkBerlalu - hadir - totalIzin);

    // Total jam mengajar (JP)
    const perJP = (typeof MENIT_PER_JP !== 'undefined') ? MENIT_PER_JP : 30;
    const laporan = getSheetData(SHEET_NAMES.LAPORAN);
    const lapUser = laporan.filter(l => l.UserID === userID && dalamBulan(tgl(l.Tanggal)));
    const totalMenit = lapUser.reduce((sum, l) => sum + hitungDurasiMenit(l.JamMulai, l.JamSelesai), 0);
    const totalJP = menitKeJP(totalMenit, perJP);
    const totalJPLabel = formatJP(totalMenit, perJP);

    return {
      success: true,
      bulan,
      tahun,
      namaBulan: bulans[bulan - 1] || '',
      hadir,
      telat,
      izin: totalIzin,
      alpha,
      totalJP,
      totalJPLabel
    };
  } catch(e) { return { success: false, message: e.message }; }
}

// ─── AMBIL DATA USER ──────────────────────────────────────

// ─── AMBIL DATA USER ──────────────────────────────────────
function getDataUser(token) {
  try {
    const session = requireUser(token);
    const userID = session.userID;
    const absensi  = getSheetData(SHEET_NAMES.ABSENSI).filter(a => a.UserID === userID).slice(-30);
    const izin     = getSheetData(SHEET_NAMES.IZIN).filter(a => a.UserID === userID).slice(-20);
    const laporan  = getSheetData(SHEET_NAMES.LAPORAN).filter(a => a.UserID === userID).slice(-20);
    const kegiatan = getSheetData(SHEET_NAMES.KEGIATAN).filter(a => a.UserID === userID).slice(-20);
    const log      = getSheetData(SHEET_NAMES.LOG).filter(a => a.UserID === userID).slice(-20);
    return { success: true, absensi, izin, laporan, kegiatan, log };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── RIWAYAT USER: PRESENSI + AKTIVITAS PER BULAN ─────────

// ─── RIWAYAT USER: PRESENSI + AKTIVITAS PER BULAN ─────────
function getRiwayatUser(token, bulan, tahun) {
  try {
    const session = requireUser(token);
    const userID = session.userID;
    const tz = 'Asia/Jakarta';
    const kunci = parseInt(tahun) + '-' + parseInt(bulan);
    const dalamBulan = v => {
      if (!v) return false;
      const d = new Date(v);
      return !isNaN(d) && Utilities.formatDate(d, tz, 'yyyy-M') === kunci;
    };
    const jam = v => v ? Utilities.formatDate(new Date(v), tz, 'HH:mm') : '';

    const presensi = getSheetData(SHEET_NAMES.ABSENSI)
      .filter(a => a.UserID === userID && dalamBulan(a.Tanggal))
      .map(a => ({
        tanggal: Utilities.formatDate(new Date(a.Tanggal), tz, 'yyyy-MM-dd'),
        masuk: jam(a.WaktuMasuk),
        pulang: jam(a.WaktuPulang),
        statusMasuk: a.StatusMasuk || '',
        statusPulang: a.StatusPulang || ''
      }))
      .sort((x, y) => y.tanggal.localeCompare(x.tanggal));

    const aktivitas = getSheetData(SHEET_NAMES.LOG)
      .filter(l => l.UserID === userID && dalamBulan(l.Waktu))
      .map(l => ({
        tanggal: Utilities.formatDate(new Date(l.Waktu), tz, 'yyyy-MM-dd'),
        jam: jam(l.Waktu),
        aksi: l.Aksi || '',
        detail: l.Detail || ''
      }))
      .reverse(); // sheet Log urut waktu naik → balik supaya terbaru di atas

    return { success: true, presensi, aktivitas };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ─── ADMIN: DASHBOARD DATA ────────────────────────────────

// ─── USER: UPDATE PROFIL SENDIRI ─────────────────────────
function updateProfile(token, data) {
  try {
    const session = requireUser(token);
    const userID  = session.userID;

    const nama      = sanitize(data.nama);
    const nip       = sanitize(data.nip);
    const mapel     = sanitize(data.mapel);
    const pwLama    = String(data.passwordLama   || '');
    const pwBaru    = String(data.passwordBaru   || '');
    const pwKonfirm = String(data.passwordKonfirm || '');

    if (!nama) return { success: false, message: 'Nama tidak boleh kosong' };

    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(SHEET_NAMES.USER);
    const rows = sh.getDataRange().getValues();

    let rowIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === userID) { rowIdx = i + 1; break; }
    }
    if (rowIdx < 0) return { success: false, message: 'User tidak ditemukan' };

    // Jika ingin ganti password, verifikasi password lama dulu
    if (pwBaru) {
      if (!pwLama) return { success: false, message: 'Masukkan password lama untuk mengganti password' };
      const storedPass = rows[rowIdx - 1][3];
      const inputHash  = hashPassword(pwLama);
      if (storedPass !== inputHash) return { success: false, message: 'Password lama salah' };
      if (pwBaru.length < 6) return { success: false, message: 'Password baru minimal 6 karakter' };
      if (pwBaru !== pwKonfirm) return { success: false, message: 'Konfirmasi password tidak cocok' };
      sh.getRange(rowIdx, 4).setValue(hashPassword(pwBaru));
    }

    sh.getRange(rowIdx, 2).setValue(nama);
    sh.getRange(rowIdx, 6).setValue(nip);
    sh.getRange(rowIdx, 7).setValue(mapel);

    const users = getSheetData(SHEET_NAMES.USER);
    const user  = users.find(u => u.ID === userID);
    addLog(userID, nama, 'UPDATE_PROFIL', 'Profil diperbarui' + (pwBaru ? ' + ganti password' : ''));

    return {
      success: true,
      message: 'Profil berhasil diperbarui!' + (pwBaru ? ' Password baru aktif.' : ''),
      user: { id: userID, nama, email: user ? user.Email : '', role: session.role, nip, mapel }
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
