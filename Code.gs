// Smart QR Attendance - Google Sheets backend (Google Apps Script)
// Paste into: Google Sheet > Extensions > Apps Script, then Deploy as a Web App.

const ADMIN_PASSWORD = 'admin';      // CHANGE THIS
const TZ = 'Asia/Kolkata';           // your timezone

const HEADERS = {
  Students:   ['id', 'name', 'email', 'course', 'password'],
  Sessions:   ['sessionId', 'subject', 'date', 'startedAt', 'expiresAt', 'status'],
  Attendance: ['timestamp', 'date', 'time', 'studentId', 'studentName', 'subject', 'sessionId', 'status']
};

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1000, HEADERS[name].length).setNumberFormat('@'); // keep everything as plain text
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const h = v.shift().map(String);
  return v.filter(r => r[0] !== '')
          .map(r => Object.fromEntries(h.map((k, i) => [k, String(r[i])])));
}

const fmt_ = (ms, f) => Utilities.formatDate(new Date(ms), TZ, f);

function doGet() { return ContentService.createTextOutput('Attendance API is running'); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let res;
  try { res = handle_(JSON.parse(e.postData.contents)); }
  catch (err) { res = { ok: false, error: String(err.message || err) }; }
  finally { lock.releaseLock(); }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(d) {
  const isAdmin = d.role === 'admin';
  let me;
  if (isAdmin) {
    if (d.id !== 'admin' || d.password !== ADMIN_PASSWORD) throw new Error('Invalid credentials!');
    me = { id: 'admin', name: 'Administrator', role: 'admin' };
  } else {
    let s = rows_('Students').find(x => x.id === String(d.id) && x.password === String(d.password));
    // Fallback for mock student credentials if not in DB
    if (!s && String(d.id) === 'student' && String(d.password) === 'student') {
      s = { id: 'student', name: 'Mock Student', role: 'student' };
    }
    if (!s) throw new Error('Invalid credentials!');
    me = { id: s.id, name: s.name, role: 'student' };
  }
  const needAdmin = () => { if (!isAdmin) throw new Error('Admin only'); };
  const now = Date.now();

  switch (d.action) {
    case 'login':
      return { ok: true, user: me };

    case 'getAll': {
      let students = rows_('Students').map(({ password, ...s }) => s);
      let attendance = rows_('Attendance');
      if (!isAdmin) { attendance = attendance.filter(a => a.studentId === me.id); students = []; }
      return { ok: true, students, sessions: rows_('Sessions'), attendance };
    }

    case 'createSession': {
      needAdmin();
      const sessionId = 'SESSION_' + now;
      const expiresAt = now + Number(d.minutes) * 60000;
      sheet_('Sessions').appendRow([sessionId, d.subject, fmt_(now, 'yyyy-MM-dd'), now, expiresAt, 'active']);
      return { ok: true, session: { sessionId, subject: d.subject, expiresAt } };
    }

    case 'endSession': {
      needAdmin();
      const sh = sheet_('Sessions');
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat().map(String);
      const i = ids.indexOf(d.sessionId);
      if (i > 0) sh.getRange(i + 1, 6).setValue('ended');
      return { ok: true };
    }

    case 'mark': {
      if (isAdmin) throw new Error('Students only');
      const s = rows_('Sessions').find(x => x.sessionId === d.sessionId);
      if (!s) throw new Error('Unrecognized QR Code');
      if (s.status !== 'active' || now > Number(s.expiresAt)) throw new Error('This QR Code has expired.');
      const already = rows_('Attendance').some(a => a.sessionId === s.sessionId && a.studentId === me.id);
      const time = fmt_(now, 'HH:mm');
      if (!already) sheet_('Attendance').appendRow([now, s.date, time, me.id, me.name, s.subject, s.sessionId, 'Present']);
      return { ok: true, already, subject: s.subject, time };
    }

    case 'addStudent': {
      needAdmin();
      const st = d.student;
      if (rows_('Students').some(x => x.id === String(st.id))) throw new Error('Student ID already exists!');
      sheet_('Students').appendRow([st.id, st.name, st.email, st.course, st.studentPassword]);
      return { ok: true };
    }

    case 'deleteStudent': {
      needAdmin();
      const sh = sheet_('Students');
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat().map(String);
      const i = ids.indexOf(String(d.studentId));
      if (i > 0) sh.deleteRow(i + 1);
      return { ok: true };
    }

    default:
      throw new Error('Unknown action');
  }
}
