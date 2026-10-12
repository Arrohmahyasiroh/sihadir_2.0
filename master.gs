// ============================================================
//  Master.gs — Master Data: Kelas & Mata Pelajaran
//  Dikelola admin (menu Master Data). Isinya otomatis menjadi
//  pilihan dropdown di form Laporan Mengajar milik guru.
// ============================================================

const MASTER_JENIS = ['Kelas', 'Mapel', 'Kegiatan'];
const MASTER_LABEL = { Kelas: 'Kelas', Mapel: 'Mata pelajaran', Kegiatan: 'Kegiatan' };
const MASTER_KELAS_AWAL = ['1A','1B','1C','1D','1E','2A','2B','2C','2D','2E'];

/** Pastikan sheet MasterData ada. Bila baru dibuat, diisi kelas awal (1A-E, 2A-E). */
function pastikanSheetMaster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.MASTER);
  if (sh) return sh;
  try {
    sh = ss.insertSheet(SHEET_NAMES.MASTER);
    sh.appendRow(['ID', 'Jenis', 'Nama', 'CreatedAt']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a237e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    MASTER_KELAS_AWAL.forEach(k => tulisBarisMaster_(sh, 'Kelas', k));
  } catch (e) {
    sh = ss.getSheetByName(SHEET_NAMES.MASTER); // sudah dibuat proses lain di saat bersamaan
    if (!sh) throw e;
  }
  return sh;
}

/** Tulis satu baris; kolom Nama dipaksa TEKS supaya "1E" tidak dibaca sebagai angka & "=..." tidak jadi rumus. */
function tulisBarisMaster_(sh, jenis, nama) {
  const baris = sh.getLastRow() + 1;
  sh.getRange(baris, 3).setNumberFormat('@');
  sh.getRange(baris, 1, 1, 4).setValues([['MST_' + Utilities.getUuid().slice(0, 8), jenis, nama, new Date()]]);
}

/** Bersihkan nama master. Tanda kutip tunggal dipertahankan (mis. "Qur'an Hadis"). */
function bersihkanNama_(v) {
  return String(v == null ? '' : v).replace(/[<>"`]/g, '').replace(/\s+/g, ' ').trim();
}

/** Kelas: urut angka lalu huruf (2A sebelum 10A). Mapel: urut abjad. */
function urutKelas_(a, b) {
  const pa = String(a.nama).match(/^(\d+)\s*(.*)$/), pb = String(b.nama).match(/^(\d+)\s*(.*)$/);
  if (pa && pb) return (parseInt(pa[1]) - parseInt(pb[1])) || pa[2].localeCompare(pb[2], 'id');
  if (pa) return -1;
  if (pb) return 1;
  return String(a.nama).localeCompare(String(b.nama), 'id');
}

function getMasterList_(jenis) {
  pastikanSheetMaster_();
  return getSheetData(SHEET_NAMES.MASTER)
    .filter(r => r.Jenis === jenis && String(r.Nama).trim() !== '')
    .map(r => ({ id: r.ID, nama: String(r.Nama).trim() }))
    .sort(jenis === 'Kelas' ? urutKelas_ : (a, b) => a.nama.localeCompare(b.nama, 'id'));
}

/**
 * Cocokkan isian guru dengan master.
 *  - Daftar master kosong  → isian bebas diterima (agar guru tidak terkunci sebelum admin mengisi).
 *  - Daftar ada isinya     → harus salah satu dari daftar (huruf besar/kecil diabaikan); yang disimpan ejaan master.
 */
function cocokkanMaster_(jenis, nilai, label) {
  const teks = String(nilai == null ? '' : nilai).trim();
  if (!teks) return { ok: false, pesan: label + ' wajib dipilih' };
  const daftar = getMasterList_(jenis);
  if (!daftar.length) return { ok: true, nilai: sanitize(teks) };
  const cocok = daftar.find(x => x.nama.toLowerCase() === teks.toLowerCase());
  if (!cocok) return { ok: false, pesan: `${label} "${teks}" tidak ada di daftar. Pilih dari pilihan yang tersedia.` };
  return { ok: true, nilai: cocok.nama };
}

// ─── GURU: PILIHAN DROPDOWN FORM LAPORAN ──────────────────
function getOpsiLaporan(token) {
  try {
    requireUser(token);
    return {
      success: true,
      kelas: getMasterList_('Kelas').map(x => x.nama),
      mapel: getMasterList_('Mapel').map(x => x.nama)
    };
  } catch (e) { return { success: false, message: e.message }; }
}

// ─── ADMIN: KELOLA MASTER ─────────────────────────────────
function getMasterData(token) {
  try {
    requireAdmin(token);
    return { success: true, kelas: getMasterList_('Kelas'), mapel: getMasterList_('Mapel'), kegiatan: getMasterList_('Kegiatan') };
  } catch (e) { return { success: false, message: e.message }; }
}

/** Jalankan fn di dalam kunci skrip, supaya klik ganda tidak menghasilkan data ganda. */
function denganKunciMaster_(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return { success: false, message: 'Sistem sedang sibuk. Coba lagi sebentar.' }; }
  try { return fn(); } finally { lock.releaseLock(); }
}

function tambahMaster(token, jenis, nama) {
  try {
    const session = requireAdmin(token);
    if (!MASTER_JENIS.includes(jenis)) return { success: false, message: 'Jenis data tidak valid' };
    nama = bersihkanNama_(nama);
    if (!nama) return { success: false, message: 'Nama wajib diisi' };
    if (nama.length > 50) return { success: false, message: 'Nama maksimal 50 karakter' };
    return denganKunciMaster_(() => {
      const sh = pastikanSheetMaster_();
      if (getMasterList_(jenis).some(x => x.nama.toLowerCase() === nama.toLowerCase())) {
        return { success: false, message: `${MASTER_LABEL[jenis]} "${nama}" sudah ada` };
      }
      tulisBarisMaster_(sh, jenis, nama);
      addLog(session.userID, 'Admin', 'MASTER_TAMBAH', `${jenis}: ${nama}`);
      return { success: true, message: `${MASTER_LABEL[jenis]} "${nama}" berhasil ditambahkan` };
    });
  } catch (e) { return { success: false, message: e.message }; }
}

function ubahMaster(token, id, nama) {
  try {
    const session = requireAdmin(token);
    id = sanitize(id);
    nama = bersihkanNama_(nama);
    if (!nama) return { success: false, message: 'Nama wajib diisi' };
    if (nama.length > 50) return { success: false, message: 'Nama maksimal 50 karakter' };
    return denganKunciMaster_(() => {
      const sh = pastikanSheetMaster_();
      const rows = sh.getDataRange().getValues();
      let idx = -1;
      for (let i = 1; i < rows.length; i++) { if (rows[i][0] === id) { idx = i; break; } }
      if (idx < 0) return { success: false, message: 'Data tidak ditemukan (mungkin sudah dihapus)' };
      const jenis = rows[idx][1], lama = String(rows[idx][2]);
      if (getMasterList_(jenis).some(x => x.id !== id && x.nama.toLowerCase() === nama.toLowerCase())) {
        return { success: false, message: `${MASTER_LABEL[jenis]} "${nama}" sudah ada` };
      }
      sh.getRange(idx + 1, 3).setNumberFormat('@').setValue(nama);
      addLog(session.userID, 'Admin', 'MASTER_UBAH', `${jenis}: ${lama} → ${nama}`);
      return { success: true, message: `Diubah dari "${lama}" menjadi "${nama}". Laporan yang sudah tersimpan tidak berubah.` };
    });
  } catch (e) { return { success: false, message: e.message }; }
}

function hapusMaster(token, id) {
  try {
    const session = requireAdmin(token);
    id = sanitize(id);
    return denganKunciMaster_(() => {
      const sh = pastikanSheetMaster_();
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === id) {
          const jenis = rows[i][1], nama = String(rows[i][2]);
          sh.deleteRow(i + 1);
          addLog(session.userID, 'Admin', 'MASTER_HAPUS', `${jenis}: ${nama}`);
          return { success: true, message: `${MASTER_LABEL[jenis]} "${nama}" dihapus. Laporan yang sudah tersimpan tidak berubah.` };
        }
      }
      return { success: false, message: 'Data tidak ditemukan (mungkin sudah dihapus)' };
    });
  } catch (e) { return { success: false, message: e.message }; }
}
