// ============================================================
//  Auth.gs — login, session, rate limiting, hashing & sanitasi
// ============================================================

// ============================================================
//  [KEAMANAN] HASHING & SANITASI
// ============================================================

// ─── FLAG HASH PASSWORD ───────────────────────────────────
// Ubah ke true untuk mengaktifkan kembali hashing SHA-256.
// Saat ini false → password disimpan & dibandingkan sebagai plaintext.
const USE_PASSWORD_HASH = false;

function hashPassword(plain) {
  if (!USE_PASSWORD_HASH) return String(plain || '');
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Sanitasi input: hapus karakter berbahaya untuk mencegah injection.
 */

function sanitize(val) {
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/[<>"'`;]/g, '')   // hapus karakter HTML/script berbahaya
    .trim()
    .slice(0, 500);             // batasi panjang
}

/**
 * Sanitasi khusus koordinat & angka.
 * Mempertahankan tanda minus di awal (untuk koordinat negatif seperti -7.250445).
 * JANGAN pakai sanitize() biasa untuk koordinat karena akan menghapus karakter '-'.
 */

function sanitizeNumber(val) {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  // Izinkan: opsional minus di awal, lalu digit dan satu titik desimal
  const match = s.match(/^-?\d*\.?\d*/);
  return match ? match[0] : '';
}

// ============================================================
//  [KEAMANAN] SESSION TOKEN
// ============================================================

/**
 * Buat token sesi acak untuk user yang baru login.
 * Token disimpan di sheet Sessions dengan waktu expired (8 jam).
 */
function createSession(userID, role) {
  const token = Utilities.getUuid();
  const expired = new Date(Date.now() + 8 * 60 * 60 * 1000); // 8 jam
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.SESSION);
  if (!sh) return null;
  // Hapus sesi lama milik user ini
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] === userID) sh.deleteRow(i + 1);
  }
  sh.appendRow([token, userID, role, new Date(), expired]);
  return token;
}

/**
 * Validasi token: kembalikan { userID, role } atau null jika tidak valid.
 */

function validateSession(token) {
  if (!token) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.SESSION);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === token) {
      const expired = new Date(rows[i][4]);
      if (now > expired) {
        sh.deleteRow(i + 1); // hapus token expired
        return null;
      }
      return { userID: rows[i][1], role: rows[i][2] };
    }
  }
  return null;
}

/**
 * Hapus sesi (logout).
 */

function destroySession(token) {
  if (!token) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.SESSION);
  if (!sh) return;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === token) { sh.deleteRow(i + 1); return; }
  }
}

// ============================================================
//  [KEAMANAN] RATE LIMITING LOGIN (lockout 5x salah)
// ============================================================
const MAX_FAIL = 5;
const LOCKOUT_MINUTES = 15;

function checkLoginAllowed(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.LOGIN_FAIL);
  if (!sh) return { allowed: true };
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === email) {
      const count = rows[i][1];
      const lastFail = new Date(rows[i][2]);
      const minutesAgo = (Date.now() - lastFail) / 60000;
      if (count >= MAX_FAIL && minutesAgo < LOCKOUT_MINUTES) {
        const sisaMenit = Math.ceil(LOCKOUT_MINUTES - minutesAgo);
        return { allowed: false, message: `Akun dikunci sementara. Coba lagi dalam ${sisaMenit} menit.` };
      }
      // Reset otomatis setelah LOCKOUT_MINUTES berlalu
      if (minutesAgo >= LOCKOUT_MINUTES) {
        sh.getRange(i + 1, 2).setValue(0);
      }
      return { allowed: true };
    }
  }
  return { allowed: true };
}

function recordLoginFail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.LOGIN_FAIL);
  if (!sh) return;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === email) {
      sh.getRange(i + 1, 2).setValue((rows[i][1] || 0) + 1);
      sh.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  sh.appendRow([email, 1, new Date()]);
}

function resetLoginFail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.LOGIN_FAIL);
  if (!sh) return;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === email) {
      sh.getRange(i + 1, 2).setValue(0);
      return;
    }
  }
}

// ============================================================
//  [KEAMANAN] GUARD ADMIN — wajib dipanggil di tiap fungsi admin
// ============================================================
function requireAdmin(token) {
  const session = validateSession(token);
  if (!session || session.role !== 'admin') {
    throw new Error('Akses ditolak: bukan admin atau sesi tidak valid');
  }
  return session;
}

function requireUser(token) {
  const session = validateSession(token);
  if (!session) {
    throw new Error('Sesi tidak valid atau sudah berakhir, silakan login kembali');
  }
  return session;
}

// ─── SETUP DATABASE ───────────────────────────────────────

// ─── LOGIN ────────────────────────────────────────────────
function login(email, password) {
  try {
    email = sanitize(email).toLowerCase();
    password = String(password || '');

    // [KEAMANAN] Cek lockout
    const lockCheck = checkLoginAllowed(email);
    if (!lockCheck.allowed) return { success: false, message: lockCheck.message };

    const users = getSheetData(SHEET_NAMES.USER);
    const hashedInput = hashPassword(password);
    const user = users.find(u =>
      u.Email === email &&
      u.Password === hashedInput &&
      (u.Aktif === true || u.Aktif === 'TRUE')
    );

    if (!user) {
      recordLoginFail(email);
      // Cek lagi apakah sekarang sudah terkunci
      const failRows = getSheetData(SHEET_NAMES.LOGIN_FAIL);
      const failData = failRows.find(r => r.Email === email);
      const failCount = failData ? (failData.FailCount || 0) : 0;
      const sisaCoba = MAX_FAIL - failCount;
      const msg = sisaCoba > 0
        ? `Email atau password salah. Sisa percobaan: ${sisaCoba}`
        : `Akun dikunci ${LOCKOUT_MINUTES} menit karena terlalu banyak percobaan gagal.`;
      return { success: false, message: msg };
    }

    resetLoginFail(email);
    const token = createSession(user.ID, user.Role);
    addLog(user.ID, user.Nama, 'LOGIN', 'Login berhasil');
    return {
      success: true,
      token,
      user: { id: user.ID, nama: user.Nama, email: user.Email, role: user.Role, nip: user.NIP, mapel: user.MapelJabatan }
    };
  } catch(e) {
    return { success: false, message: 'Terjadi kesalahan: ' + e.message };
  }
}

// ─── LOGOUT ───────────────────────────────────────────────

// ─── LOGOUT ───────────────────────────────────────────────
function logout(token) {
  destroySession(token);
  return { success: true };
}

// ─── ABSENSI MASUK ────────────────────────────────────────
