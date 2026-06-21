/**
 * TJKT CLASS PLACEMENT SYSTEM - SMK NEGERI 1 BUKATEJA
 * Backend: Google Apps Script (Web App) + Google Sheets as DB
 *
 * SHEETS USED (auto-created by setupDatabase()):
 *  - DataSiswa   : NISN | Nama | Password | Role | KelasAsal | Status
 *  - Jawaban     : NISN | Nama | Timestamp | Q1..Q36 | Pilihan1 | Pilihan2 | Pilihan3
 *                  | SkorFiber | SkorNetwork | SkorCloud | Prioritas1 | Prioritas2 | Prioritas3
 *                  | RekomendasiAwal | KelasFinal | Status
 *  - Settings    : Key | Value   (quota config etc.)
 *  - ActivityLog : Timestamp | NISN | Action | Detail
 */

// ====================== CONFIG ======================
const SHEET_SISWA   = 'DataSiswa';
const SHEET_JAWABAN = 'Jawaban';
const SHEET_SETTING = 'Settings';
const SHEET_LOG     = 'ActivityLog';

const DEFAULT_QUOTA = {
  Fiber:   { min: 32, max: 33, label: 'Fiber Optic Technician (Kelas A)' },
  Network: { min: 43, max: 44, label: 'Network Engineer (Kelas B)' },
  Cloud:   { min: 32, max: 33, label: 'Cloud Engineer (Kelas C)' }
};

const SKOR_GROUP = {
  Fiber:   [1, 2, 3, 11, 21, 25, 28, 29, 32],
  Network: [4, 5, 6, 12, 13, 14, 19, 22, 27, 30, 35],
  Cloud:   [7, 8, 10, 15, 16, 17, 18, 20, 23, 24, 31, 33, 34, 36]
};

// ====================== SPREADSHEET BINDING ======================
/**
 * This project is standalone (not bound to a Sheet), so we create/open
 * our own database Spreadsheet and remember its ID in Script Properties.
 */
function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DB_SHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // stored id invalid/deleted, fall through to recreate
    }
  }
  const ss = SpreadsheetApp.create('DB - TJKT Class Placement');
  props.setProperty('DB_SHEET_ID', ss.getId());
  return ss;
}

// ====================== WEB APP ENTRY ======================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('TJKT Class Placement - SMK Negeri 1 Bukateja')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================== SETUP / INSTALL ======================
/**
 * Run this ONCE manually from the Apps Script editor to create the
 * spreadsheet structure. Re-running is safe (won't duplicate headers).
 */
function setupDatabase() {
  const ss = getSS_();

  // DataSiswa
  let sh = ss.getSheetByName(SHEET_SISWA);
  if (!sh) sh = ss.insertSheet(SHEET_SISWA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['NISN', 'Nama', 'Password', 'Role', 'KelasAsal', 'Status']);
    sh.appendRow(['0072345612', 'Adi Prasetyo', 'tjkt2024', 'siswa', 'IX-A', 'aktif']);
    sh.appendRow(['admin', 'Administrator', 'admin123', 'admin', '-', 'aktif']);
    sh.setFrozenRows(1);
  }

  // Jawaban
  sh = ss.getSheetByName(SHEET_JAWABAN);
  if (!sh) sh = ss.insertSheet(SHEET_JAWABAN);
  if (sh.getLastRow() === 0) {
    const qCols = [];
    for (let i = 1; i <= 36; i++) qCols.push('Q' + i);
    const header = ['NISN', 'Nama', 'Timestamp']
      .concat(qCols)
      .concat(['Pilihan1', 'Pilihan2', 'Pilihan3',
        'SkorFiber', 'SkorNetwork', 'SkorCloud',
        'RekomendasiAwal', 'KelasFinal', 'Status']);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }

  // Settings
  sh = ss.getSheetByName(SHEET_SETTING);
  if (!sh) sh = ss.insertSheet(SHEET_SETTING);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['QuotaFiberMax', 33]);
    sh.appendRow(['QuotaNetworkMax', 44]);
    sh.appendRow(['QuotaCloudMax', 33]);
    sh.setFrozenRows(1);
  }

  // ActivityLog
  sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) sh = ss.insertSheet(SHEET_LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'NISN', 'Action', 'Detail']);
    sh.setFrozenRows(1);
  }

  Logger.log('Database berhasil disiapkan! Spreadsheet ID: ' + ss.getId());
  Logger.log('Buka di: ' + ss.getUrl());
}

// ====================== HELPERS ======================
function getSheet_(name) {
  return getSS_().getSheetByName(name);
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data
    .filter(row => row.join('') !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function logActivity_(nisn, action, detail) {
  const sh = getSheet_(SHEET_LOG);
  sh.appendRow([new Date(), nisn, action, detail || '']);
}

function getQuota_() {
  const sh = getSheet_(SHEET_SETTING);
  const rows = sheetToObjects_(sh);
  const map = {};
  rows.forEach(r => map[r.Key] = r.Value);
  return {
    Fiber:   { max: Number(map.QuotaFiberMax)   || DEFAULT_QUOTA.Fiber.max },
    Network: { max: Number(map.QuotaNetworkMax) || DEFAULT_QUOTA.Network.max },
    Cloud:   { max: Number(map.QuotaCloudMax)   || DEFAULT_QUOTA.Cloud.max }
  };
}

// ====================== AUTH ======================
function login(nisn, password, role) {
  nisn = String(nisn).trim();
  const sh = getSheet_(SHEET_SISWA);
  const rows = sheetToObjects_(sh);
  const user = rows.find(r => String(r.NISN).trim() === nisn && String(r.Password) === String(password));

  if (!user) {
    return { success: false, message: 'NISN atau password salah.' };
  }
  if (role && String(user.Role).toLowerCase() !== role.toLowerCase()) {
    return { success: false, message: 'Akun ini tidak terdaftar sebagai ' + role + '.' };
  }

  logActivity_(nisn, 'LOGIN', user.Role);

  return {
    success: true,
    user: {
      nisn: user.NISN,
      nama: user.Nama,
      role: user.Role,
      kelasAsal: user.KelasAsal
    }
  };
}

// ====================== SISWA: STATUS & SURVEY ======================
function getStudentStatus(nisn) {
  nisn = String(nisn).trim();
  const sh = getSheet_(SHEET_JAWABAN);
  const rows = sheetToObjects_(sh);
  const existing = rows.find(r => String(r.NISN).trim() === nisn);
  if (!existing) {
    return { submitted: false };
  }
  return {
    submitted: true,
    result: {
      skorFiber: existing.SkorFiber,
      skorNetwork: existing.SkorNetwork,
      skorCloud: existing.SkorCloud,
      rekomendasiAwal: existing.RekomendasiAwal,
      kelasFinal: existing.KelasFinal,
      status: existing.Status,
      timestamp: existing.Timestamp
    }
  };
}

/**
 * payload = {
 *   nisn, nama,
 *   answers: { '1': 3, '2': 4, ... '36': 2 },
 *   choices: { p1:'Fiber'|'Network'|'Cloud', p2:..., p3:... }
 * }
 */
function submitAngket(payload) {
  const nisn = String(payload.nisn).trim();
  const sh = getSheet_(SHEET_JAWABAN);
  const rows = sheetToObjects_(sh);
  const already = rows.find(r => String(r.NISN).trim() === nisn);
  if (already) {
    return { success: false, message: 'Anda sudah mengisi angket sebelumnya.' };
  }

  const answers = payload.answers;
  let skorFiber = 0, skorNetwork = 0, skorCloud = 0;
  SKOR_GROUP.Fiber.forEach(q => skorFiber += Number(answers[q] || 0));
  SKOR_GROUP.Network.forEach(q => skorNetwork += Number(answers[q] || 0));
  SKOR_GROUP.Cloud.forEach(q => skorCloud += Number(answers[q] || 0));

  const scores = { Fiber: skorFiber, Network: skorNetwork, Cloud: skorCloud };
  const rekomendasiAwal = Object.keys(scores).reduce((a, b) => scores[a] >= scores[b] ? a : b);

  // Build row
  const qCols = [];
  for (let i = 1; i <= 36; i++) qCols.push(Number(answers[i] || 0));

  const rowValues = [nisn, payload.nama, new Date()]
    .concat(qCols)
    .concat([payload.choices.p1, payload.choices.p2, payload.choices.p3])
    .concat([skorFiber, skorNetwork, skorCloud, rekomendasiAwal, '', 'menunggu penempatan']);

  sh.appendRow(rowValues);
  logActivity_(nisn, 'SUBMIT_ANGKET', rekomendasiAwal);

  // Real-time placement
  const finalClass = assignClassRealtime_(nisn);

  return {
    success: true,
    result: {
      skorFiber, skorNetwork, skorCloud,
      rekomendasiAwal,
      kelasFinal: finalClass
    }
  };
}

// ====================== CLASS ASSIGNMENT (REAL-TIME) ======================
/**
 * Called right after a student submits. Assigns a class immediately based on
 * current quota availability, ranked by score (highest score for that track
 * gets priority if quota is contested). If preferred track is full, falls
 * back to the next-highest-scoring track that still has room.
 */
function assignClassRealtime_(nisn) {
  const sh = getSheet_(SHEET_JAWABAN);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  const quota = getQuota_();
  const counts = { Fiber: 0, Network: 0, Cloud: 0 };
  for (let i = 1; i < data.length; i++) {
    const kf = data[i][col('KelasFinal')];
    if (kf) counts[kf] = (counts[kf] || 0) + 1;
  }

  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('NISN')]).trim() === String(nisn).trim()) { targetRow = i; break; }
  }
  if (targetRow === -1) return '';

  const scores = {
    Fiber: Number(data[targetRow][col('SkorFiber')]),
    Network: Number(data[targetRow][col('SkorNetwork')]),
    Cloud: Number(data[targetRow][col('SkorCloud')])
  };
  const ranked = Object.keys(scores).sort((a, b) => scores[b] - scores[a]);

  let finalClass = null;
  for (const track of ranked) {
    if (counts[track] < quota[track].max) {
      finalClass = track;
      break;
    }
  }
  if (!finalClass) {
    // All full (shouldn't normally happen) - place in track with smallest overflow
    finalClass = ranked[0];
  }

  sh.getRange(targetRow + 1, col('KelasFinal') + 1).setValue(finalClass);
  sh.getRange(targetRow + 1, col('Status') + 1).setValue('terplaced (otomatis)');

  return finalClass;
}

/**
 * Admin tool: re-balance ALL placements from scratch using full ranking
 * (used by "Sync / Recalculate" button for fairness across all students).
 */
function rebalanceAllPlacements() {
  const sh = getSheet_(SHEET_JAWABAN);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const quota = getQuota_();

  const students = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][col('NISN')]) continue;
    students.push({
      row: i,
      nisn: data[i][col('NISN')],
      scores: {
        Fiber: Number(data[i][col('SkorFiber')]),
        Network: Number(data[i][col('SkorNetwork')]),
        Cloud: Number(data[i][col('SkorCloud')])
      }
    });
  }

  const counts = { Fiber: 0, Network: 0, Cloud: 0 };
  const tracks = ['Fiber', 'Network', 'Cloud'];

  // Sort by each track's score descending, assign greedily with global pass
  // Strategy: for each track, sort students by that track's score desc,
  // then iterate all students sorted by their MAX score desc to prioritize strongest fits first.
  students.sort((a, b) => {
    const maxA = Math.max(a.scores.Fiber, a.scores.Network, a.scores.Cloud);
    const maxB = Math.max(b.scores.Fiber, b.scores.Network, b.scores.Cloud);
    return maxB - maxA;
  });

  students.forEach(s => {
    const ranked = tracks.slice().sort((a, b) => s.scores[b] - s.scores[a]);
    let assigned = null;
    for (const t of ranked) {
      if (counts[t] < quota[t].max) { assigned = t; break; }
    }
    if (!assigned) assigned = ranked[0];
    counts[assigned]++;
    sh.getRange(s.row + 1, col('KelasFinal') + 1).setValue(assigned);
    sh.getRange(s.row + 1, col('Status') + 1).setValue('terplaced (sync admin)');
  });

  logActivity_('admin', 'REBALANCE_ALL', JSON.stringify(counts));
  return { success: true, counts };
}

// ====================== ADMIN ======================
function getAdminStats() {
  const totalStudents = sheetToObjects_(getSheet_(SHEET_SISWA)).filter(r => r.Role === 'siswa').length;
  const jawaban = sheetToObjects_(getSheet_(SHEET_JAWABAN));
  const completed = jawaban.length;
  const counts = { Fiber: 0, Network: 0, Cloud: 0 };
  let pending = 0;
  jawaban.forEach(r => {
    if (r.KelasFinal) counts[r.KelasFinal] = (counts[r.KelasFinal] || 0) + 1;
    else pending++;
  });
  const quota = getQuota_();

  return {
    totalStudents,
    completed,
    pending,
    distribution: counts,
    quota,
    participationRate: totalStudents ? Math.round((completed / totalStudents) * 1000) / 10 : 0
  };
}

function getAllResults() {
  const jawaban = sheetToObjects_(getSheet_(SHEET_JAWABAN));
  return jawaban.map(r => ({
    nisn: r.NISN,
    nama: r.Nama,
    skorFiber: r.SkorFiber,
    skorNetwork: r.SkorNetwork,
    skorCloud: r.SkorCloud,
    rekomendasiAwal: r.RekomendasiAwal,
    kelasFinal: r.KelasFinal,
    status: r.Status,
    timestamp: r.Timestamp
  }));
}

function manualOverridePlacement(nisn, newClass) {
  const sh = getSheet_(SHEET_JAWABAN);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('NISN')]).trim() === String(nisn).trim()) {
      sh.getRange(i + 1, col('KelasFinal') + 1).setValue(newClass);
      sh.getRange(i + 1, col('Status') + 1).setValue('terplaced (override admin)');
      logActivity_(nisn, 'OVERRIDE', newClass);
      return { success: true };
    }
  }
  return { success: false, message: 'Siswa tidak ditemukan.' };
}

/**
 * Generates a fresh Google Sheet with the result matrix and returns an
 * xlsx export URL the admin can open/download directly.
 */
function exportResultsToExcel() {
  const jawaban = sheetToObjects_(getSheet_(SHEET_JAWABAN));
  const newSS = SpreadsheetApp.create('Hasil Pembagian Kelas TJKT - ' + new Date().toLocaleDateString());
  const sh = newSS.getSheets()[0];
  sh.setName('Hasil Pembagian Kelas');
  sh.appendRow(['NISN', 'Nama', 'Skor Fiber', 'Skor Network', 'Skor Cloud', 'Rekomendasi Awal', 'Kelas Final', 'Status']);
  jawaban.forEach(r => {
    sh.appendRow([r.NISN, r.Nama, r.SkorFiber, r.SkorNetwork, r.SkorCloud, r.RekomendasiAwal, r.KelasFinal, r.Status]);
  });
  sh.getRange(1, 1, 1, 8).setFontWeight('bold');
  sh.autoResizeColumns(1, 8);

  const fileId = newSS.getId();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  return { success: true, url: exportUrl };
}

function getStudentDirectory() {
  return sheetToObjects_(getSheet_(SHEET_SISWA)).filter(r => r.Role === 'siswa');
}

function addStudent(nisn, nama, password, kelasAsal) {
  const sh = getSheet_(SHEET_SISWA);
  sh.appendRow([nisn, nama, password, 'siswa', kelasAsal || '-', 'aktif']);
  return { success: true };
}
