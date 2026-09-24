/**
 * DBI Attendance — mobile web app (Google Apps Script)
 * Bound to the "DBI Attendance App Control" Google Sheet.
 * Faculty log in with Name + 4-digit PIN, mark Present/Absent for an hour,
 * and the result is written into the Hour-wise Attendance Tracker registers.
 */

// ---- register layout (must match the Hour-wise Attendance Tracker) ----
const COL = { date: 1, day: 2, per: 3, time: 4, subj: 5, fac: 6, chg: 7, abs: 8, status: 9, mby: 16, mat: 17 };
const FIRST_ROW = 5;
const LAST_COL = 17;
const MAX_FAILS = 5;          // wrong PIN attempts before a 15-minute lock

// ------------------------------------------------------------------ web entry
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DBI Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------------ helpers
function ctl_() { return SpreadsheetApp.getActive(); }

function facultyRows_() {
  const sh = ctl_().getSheetByName('Faculty');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getDisplayValues()
    .filter(r => String(r[0]).trim() !== '');
}

function getFacultyNames() {
  return facultyRows_()
    .filter(r => String(r[3]).trim().toUpperCase() !== 'NO')
    .map(r => String(r[0]).trim())
    .sort();
}

function auth_(name, pin) {
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + String(name).toLowerCase();
  const fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) throw new Error('Too many wrong PINs. Try again after 15 minutes.');
  const row = facultyRows_().find(r => String(r[0]).trim().toLowerCase() === String(name).trim().toLowerCase());
  if (!row) throw new Error('Name not found. Contact the coordinator.');
  if (String(row[3]).trim().toUpperCase() === 'NO') throw new Error('This login is disabled. Contact the coordinator.');
  if (String(row[1]).trim() !== String(pin).trim()) {
    cache.put(key, String(fails + 1), 900);
    throw new Error('Wrong PIN.');
  }
  cache.remove(key);
  const aliases = [row[0]].concat(String(row[2] || '').split(','))
    .map(s => String(s).trim().toLowerCase()).filter(Boolean);
  return { name: String(row[0]).trim(), aliases: aliases };
}

function trackers_() {
  const sh = ctl_().getSheetByName('Settings');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 2).getDisplayValues()
    .map((r, i) => ({ idx: i, year: String(r[0]).trim(), url: String(r[1]).trim() }))
    .filter(t => t.year && t.url);
}

function openTracker_(idx) {
  const t = trackers_().find(x => x.idx === Number(idx));
  if (!t) throw new Error('Tracker not configured in Settings.');
  return { t: t, ss: SpreadsheetApp.openByUrl(t.url) };
}

function dkey_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return '';
}

function todayKey_(tz) { return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); }

function registers_(ss) {
  return ss.getSheets().filter(s => s.getName().indexOf('Register ') === 0);
}

// rows of a register for today: [{row, per, time, subj, fac, chg, abs, mby}]
function todayRows_(sh, tz) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  const dates = sh.getRange(FIRST_ROW, COL.date, last - FIRST_ROW + 1, 1).getValues();
  const tk = todayKey_(tz);
  let a = -1, b = -1;
  for (let i = 0; i < dates.length; i++) {
    if (dkey_(dates[i][0], tz) === tk) { if (a < 0) a = i; b = i; }
  }
  if (a < 0) return [];
  const vals = sh.getRange(FIRST_ROW + a, 1, b - a + 1, LAST_COL).getDisplayValues();
  return vals.map((v, j) => ({
    row: FIRST_ROW + a + j, per: v[COL.per - 1], time: v[COL.time - 1], subj: v[COL.subj - 1],
    fac: v[COL.fac - 1], chg: v[COL.chg - 1], abs: String(v[COL.abs - 1]).trim(), mby: v[COL.mby - 1]
  }));
}

function statusOf_(abs) {
  if (!abs) return { marked: false, label: 'Not marked' };
  const u = abs.toUpperCase();
  if (u === 'NC' || u === 'HOLIDAY' || u === 'H') return { marked: true, label: 'Class not held' };
  if (u === 'NIL') return { marked: true, label: 'All present' };
  const n = abs.split(/[\s,;./]+/).filter(x => /^\d+$/.test(x)).length;
  return { marked: true, label: n + ' absent' };
}

function isMine_(fac, aliases) {
  const f = String(fac).toLowerCase();
  return aliases.some(a => a && f.indexOf(a) >= 0);
}

// ------------------------------------------------------------------ API: login + my hours
function login(name, pin) {
  const me = auth_(name, pin);
  const hours = [];
  const years = [];
  trackers_().forEach(t => {
    years.push({ idx: t.idx, year: t.year });
    try {
      const ss = SpreadsheetApp.openByUrl(t.url);
      const tz = ss.getSpreadsheetTimeZone();
      registers_(ss).forEach(sh => {
        const cls = sh.getName().replace('Register ', '');
        todayRows_(sh, tz).forEach(r => {
          if (isMine_(r.fac, me.aliases)) {
            const st = statusOf_(r.abs);
            hours.push({ idx: t.idx, year: t.year, sheet: sh.getName(), cls: cls, row: r.row, per: r.per,
                         time: r.time, subj: r.chg || r.subj, marked: st.marked, label: st.label });
          }
        });
      });
    } catch (e) { /* skip unreachable tracker */ }
  });
  return { name: me.name, today: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'EEE, dd MMM yyyy'),
           hours: hours, years: years };
}

// ------------------------------------------------------------------ API: substitution pickers
function getClasses(name, pin, idx) {
  auth_(name, pin);
  const o = openTracker_(idx);
  return registers_(o.ss).map(sh => ({ sheet: sh.getName(), cls: sh.getName().replace('Register ', '') }));
}

function getClassHoursToday(name, pin, idx, sheetName) {
  auth_(name, pin);
  const o = openTracker_(idx);
  const sh = o.ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Class not found.');
  const tz = o.ss.getSpreadsheetTimeZone();
  return todayRows_(sh, tz).map(r => {
    const st = statusOf_(r.abs);
    return { idx: Number(idx), year: o.t.year, sheet: sheetName, cls: sheetName.replace('Register ', ''), row: r.row,
             per: r.per, time: r.time, subj: r.chg || r.subj, fac: r.fac, marked: st.marked, label: st.label };
  });
}

// ------------------------------------------------------------------ API: open an hour for marking
function getHour(name, pin, idx, sheetName, row) {
  auth_(name, pin);
  const o = openTracker_(idx);
  const ss = o.ss, tz = ss.getSpreadsheetTimeZone();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Class not found.');
  const v = sh.getRange(row, 1, 1, LAST_COL).getValues()[0];
  const d = sh.getRange(row, 1, 1, LAST_COL).getDisplayValues()[0];
  if (dkey_(v[COL.date - 1], tz) !== todayKey_(tz)) throw new Error('Only today\'s hours can be marked. Ask the coordinator for older corrections.');
  const cls = sheetName.replace('Register ', '');
  // students: roll + name only
  const sl = ss.getSheetByName('Student List');
  const sv = sl.getRange(5, 1, Math.max(sl.getLastRow() - 4, 1), 4).getDisplayValues();
  const students = sv.filter(r => String(r[1]).trim() === cls && String(r[3]).trim() !== '')
    .map(r => ({ roll: Number(r[0]), name: String(r[3]).trim() }));
  // subjects for this class (for substitution)
  const setup = ss.getSheetByName('Setup');
  const su = setup.getRange(11, 3, 40, 2).getDisplayValues();
  const subjects = su.filter(r => String(r[0]).trim() === cls && r[1]).map(r => String(r[1]).trim());
  const abs = String(d[COL.abs - 1]).trim();
  const absent = abs.split(/[\s,;./]+/).filter(x => /^\d+$/.test(x)).map(Number);
  return {
    idx: Number(idx), year: o.t.year, sheet: sheetName, cls: cls, row: Number(row),
    date: Utilities.formatDate(v[COL.date - 1], tz, 'EEE, dd MMM yyyy'), dkey: dkey_(v[COL.date - 1], tz),
    per: d[COL.per - 1], time: d[COL.time - 1], subj: d[COL.subj - 1], fac: d[COL.fac - 1],
    chg: d[COL.chg - 1], notHeld: /^(NC|HOLIDAY|H)$/i.test(abs), marked: abs !== '',
    markedBy: d[COL.mby - 1], absent: absent, students: students, subjects: subjects
  };
}

// ------------------------------------------------------------------ API: submit
/**
 * p = {idx, sheet, row, dkey, per, notHeld, absent:[rolls], subject, substitution}
 */
function submitAttendance(name, pin, p) {
  const me = auth_(name, pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const o = openTracker_(p.idx);
    const ss = o.ss, tz = ss.getSpreadsheetTimeZone();
    const sh = ss.getSheetByName(p.sheet);
    if (!sh) throw new Error('Class not found.');
    const v = sh.getRange(p.row, 1, 1, LAST_COL).getValues()[0];
    const d = sh.getRange(p.row, 1, 1, LAST_COL).getDisplayValues()[0];
    const dk = dkey_(v[COL.date - 1], tz);
    if (dk !== todayKey_(tz)) throw new Error('Only today\'s hours can be marked.');
    if (dk !== p.dkey || String(d[COL.per - 1]) !== String(p.per)) throw new Error('The register changed. Please reopen this hour.');
    let absText;
    if (p.notHeld) absText = 'NC';
    else {
      const rolls = (p.absent || []).map(Number).filter(n => n > 0).sort((a, b) => a - b);
      absText = rolls.length ? rolls.join(', ') : 'NIL';
    }
    const ttSubj = String(d[COL.subj - 1]);
    const chg = (p.subject && p.subject !== ttSubj) ? p.subject : '';
    const now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
    sh.getRange(p.row, COL.abs).setNumberFormat('@').setValue(absText);
    sh.getRange(p.row, COL.chg).setValue(chg);
    sh.getRange(p.row, COL.mby).setValue(me.name + (p.substitution ? ' (substitution)' : ''));
    sh.getRange(p.row, COL.mat).setValue(now);
    SpreadsheetApp.flush();
    // audit log
    let log = ctl_().getSheetByName('Log');
    if (!log) {
      log = ctl_().insertSheet('Log');
      log.appendRow(['Timestamp', 'Faculty', 'Year', 'Class', 'Date', 'Hour', 'Subject', 'Substitution', 'Recorded', 'Previous entry']);
    }
    log.appendRow([now, me.name, o.t.year, p.sheet.replace('Register ', ''), dk, p.per, chg || ttSubj,
                   p.substitution ? 'Yes' : '', absText, String(d[COL.abs - 1])]);
    const total = Number(p.total || 0);
    const nAbs = p.notHeld ? 0 : (p.absent || []).length;
    return {
      ok: true,
      summary: p.notHeld ? 'Recorded as CLASS NOT HELD'
                         : (total ? (total - nAbs) + ' present, ' + nAbs + ' absent' : nAbs + ' absent'),
      cls: p.sheet.replace('Register ', ''), per: p.per, subject: chg || ttSubj, at: now
    };
  } finally {
    lock.releaseLock();
  }
}
