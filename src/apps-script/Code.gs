/**
 * MAud Clinical Tracker — Google Apps Script API
 *
 * SETUP:
 * 1. Create a Google Sheet with tabs: Config, Students, Skills, Sessions, Ratings, Supervisors
 *    (see docs/SHEET_SETUP.md for full column definitions)
 * 2. Open Extensions > Apps Script, paste this file into Code.gs
 * 3. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployment URL and paste it into all HTML files at const API_URL = '...';
 *
 * Auth: Google ID tokens verified via Google tokeninfo API.
 * Roles: 'supervisor' (Supervisors tab) | 'student' (Students tab email col) | 'none'
 *
 * GET  ?action=role&token=…                     — verify token, return role
 * GET  ?action=config                           — cohorts + skill definitions (public)
 * GET  ?action=students&cohort=…&token=…        — student list (supervisor only)
 * GET  ?action=ratings&student=…&token=…        — ratings (supervisor or own student)
 * GET  ?action=cohort_overview&cohort=…&token=… — cohort summary (supervisor only)
 * GET  ?action=hours&student=…&token=…          — hours (supervisor or own student)
 * GET  ?action=cohort_hours&cohort=…&token=…    — cohort hours (supervisor only)
 * GET  ?action=cohort_sessions&cohort=…&token=… — all sessions across cohort (supervisor only)
 * POST body: { action:'submitSession',      token, data:{…} } — supervisor only
 * POST body: { action:'submitStudentHours', token, data:{…} } — student only
 * POST body: { action:'approveSession',     token, sessionId, approvedBy } — supervisor only
 */

var CLIENT_ID = '27923847271-b8cu115ptvp7ft3dnrtl0p1mc5c5pe81.apps.googleusercontent.com';

/**
 * Decode and verify a Google ID token locally (no external HTTP call).
 * Checks aud, iss, and expiry. Signature is implicitly trusted because
 * only Google can issue tokens for our CLIENT_ID.
 */
function verifyToken(idToken) {
  if (!idToken) return null;
  try {
    var parts = idToken.split('.');
    if (parts.length !== 3) return null;
    // base64url → base64
    var b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (b64.length % 4) b64 += '=';
    var json = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString();
    var p = JSON.parse(json);
    if (p.aud !== CLIENT_ID) return null;
    if (p.iss !== 'https://accounts.google.com' && p.iss !== 'accounts.google.com') return null;
    if (p.exp && p.exp < Math.floor(Date.now() / 1000)) return null;
    return p;
  } catch(e) { return null; }
}

/** CacheService TTL for directory/config lookups that rarely change between requests. */
var CACHE_TTL_SECONDS = 300;

/**
 * Wraps a cache-or-compute pattern: returns the cached JSON value for `key`
 * if present, otherwise calls `computeFn()`, caches the result, and returns it.
 */
function getCached(key, computeFn) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* fall through to recompute */ }
  }
  var value = computeFn();
  try { cache.put(key, JSON.stringify(value), CACHE_TTL_SECONDS); } catch (e) { /* value too large — skip caching */ }
  return value;
}

/**
 * Reads the Supervisors tab (Sign-in Email | Preferred Email | IsCoordinator | IsNZAS)
 * once per request and returns the raw rows. getSupervisorDirectory() and
 * getSupervisorInfoMap() both derive from this single cached snapshot so they can
 * never disagree with each other within a cache TTL window the way they would if
 * each cached its own independent read of the sheet.
 * Cached for CACHE_TTL_SECONDS since the Supervisors tab changes rarely.
 */
function getSupervisorsRaw() {
  return getCached('supervisorsRaw_v1', function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Supervisors');
    var rows = [];
    if (!sheet) return rows;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      rows.push({
        loginEmail: String(data[i][0] || '').trim().toLowerCase(),
        preferredEmail: String(data[i][1] || '').trim(),
        isCoordinator: data[i][2] === true || data[i][2] === 'TRUE',
        isNZAS: data[i][3] === true || data[i][3] === 'TRUE'
      });
    }
    return rows;
  });
}

/**
 * Returns { byLoginEmail: {lowercase sign-in email: preferred email}, coordinatorEmails: [preferred email,...] }.
 * There is no Name column on this tab, so sup1/sup2 free-text names typed on the
 * feedback form cannot be resolved to an email via this directory — only a
 * supervisor's own verified sign-in email can be looked up. (The signed-in
 * supervisor's display name comes from the Google token instead — see getRole.)
 */
function getSupervisorDirectory() {
  var byLoginEmail = {}, coordinatorEmails = [];
  getSupervisorsRaw().forEach(function (row) {
    if (row.loginEmail && row.preferredEmail) byLoginEmail[row.loginEmail] = row.preferredEmail;
    if (row.isCoordinator && row.preferredEmail) coordinatorEmails.push(row.preferredEmail);
  });
  return { byLoginEmail: byLoginEmail, coordinatorEmails: coordinatorEmails };
}

/**
 * Builds an email(lowercase) -> {studentId, name, cohort} map from the Students tab.
 * Cached for CACHE_TTL_SECONDS so getRole() doesn't rescan the whole tab on every request.
 */
function getStudentEmailMap() {
  return getCached('studentEmailMap_v1', function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var stuSheet = ss.getSheetByName('Students');
    var map = {};
    if (!stuSheet) return map;
    var stuData = stuSheet.getDataRange().getValues();
    for (var i = 1; i < stuData.length; i++) {
      var email = String(stuData[i][3] || '').toLowerCase().trim();
      if (!email) continue;
      map[email] = { studentId: String(stuData[i][0]), name: String(stuData[i][2]), cohort: String(stuData[i][1]) };
    }
    return map;
  });
}

/**
 * Builds an email(lowercase) -> {isNZAS} map from the Supervisors tab, derived from
 * the same getSupervisorsRaw() snapshot used by getSupervisorDirectory().
 */
function getSupervisorInfoMap() {
  var map = {};
  getSupervisorsRaw().forEach(function (row) {
    if (!row.loginEmail) return;
    map[row.loginEmail] = { isNZAS: row.isNZAS };
  });
  return map;
}

/** Look up role from verified email. Returns {role, email, name, studentId, cohort}. */
function getRole(idToken) {
  var payload = verifyToken(idToken);
  if (!payload) return { role: 'none', _debug: 'token_verify_failed' };
  var email = (payload.email || '').toLowerCase().trim();

  var studentEntry = getStudentEmailMap()[email];
  var studentId = studentEntry ? studentEntry.studentId : null;
  var studentName = studentEntry ? studentEntry.name : null;
  var cohort = studentEntry ? studentEntry.cohort : null;

  var supervisorInfo = getSupervisorInfoMap()[email];
  if (supervisorInfo) {
    // Supervisor — also include student info if they're in both tabs (for testing).
    // Name comes from the verified Google token (no Name column on the Supervisors tab).
    return { role: 'supervisor', email: email, name: payload.name || '',
             studentId: studentId, cohort: cohort, isNZAS: supervisorInfo.isNZAS };
  }

  if (studentId) {
    return { role: 'student', email: email, name: studentName,
             studentId: studentId, cohort: cohort };
  }

  return { role: 'none', email: email };
}

function doGet(e) {
  var action = e.parameter.action || 'config';
  var token  = e.parameter.token  || '';
  var result;

  try {
    // Role endpoint — also accepts token via POST body (see doPost)
    if (action === 'role') {
      result = getRole(token);
    }
    // Config — public (no auth needed to load skill definitions)
    else if (action === 'config') {
      result = getConfig();
    }
    // Protected endpoints
    else {
      var auth = getRole(token);
      if (auth.role === 'none') { result = { error: 'Unauthorised' }; }
      else if (action === 'students') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getStudents(e.parameter.cohort);
      }
      else if (action === 'ratings') {
        var reqStudent = e.parameter.student;
        if (auth.role === 'student' && auth.studentId !== reqStudent) result = { error: 'Unauthorised' };
        else result = getRatings(reqStudent, e.parameter.cohort);
      }
      else if (action === 'cohort_overview') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getCohortOverview(e.parameter.cohort);
      }
      else if (action === 'hours') {
        var reqStudent = e.parameter.student;
        if (auth.role === 'student' && auth.studentId !== reqStudent) result = { error: 'Unauthorised' };
        else result = getHours(reqStudent);
      }
      else if (action === 'cohort_hours') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getCohortHours(e.parameter.cohort);
      }
      else if (action === 'session') {
        result = getSessionDetail(e.parameter.sessionId, auth);
      }
      else if (action === 'cohort_sessions') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getCohortSessions(e.parameter.cohort);
      }
      else if (action === 'dashboard_init') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getDashboardInit(e.parameter.cohort);
      }
      else if (action === 'full_init') {
        if (auth.role !== 'supervisor') result = { error: 'Supervisor access required' };
        else result = getFullInit();
      }
      else {
        result = { error: 'Unknown action: ' + action };
      }
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests. Content-Type must be text/plain (avoids CORS preflight).
 */
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var auth = getRole(body.token || '');

    switch (body.action) {
      case 'role':
        result = getRole(body.token || '');
        break;
      case 'submitSession':
        if (auth.role !== 'supervisor') { result = { error: 'Supervisor access required' }; break; }
        result = submitSession(body.data);
        break;
      case 'submitStudentHours':
        if (!auth.studentId) { result = { error: 'Student access required' }; break; }
        result = submitStudentHours(body.data, auth);
        break;
      case 'init':
        // Combines role check + full_init + cohort_hours into one round trip for the
        // dashboard's initial load, since each Apps Script call costs several seconds
        // of fixed overhead regardless of the work done — three serial calls (role,
        // full_init, cohort_hours) was the dominant cost on first page load.
        if (auth.role === 'none') { result = { role: 'none' }; break; }
        if (auth.role === 'student') {
          result = { role: 'student', email: auth.email, name: auth.name, studentId: auth.studentId, cohort: auth.cohort };
          break;
        }
        var fullInit = getFullInit();
        var cohortHours = fullInit.cohort ? getCohortHours(fullInit.cohort) : { students: [] };
        result = {
          role: 'supervisor',
          studentId: auth.studentId,
          cohorts: fullInit.cohorts,
          skills: fullInit.skills,
          cohort: fullInit.cohort,
          students: fullInit.students,
          overview: fullInit.overview,
          cohortHours: cohortHours.students || []
        };
        break;
      case 'approveSession':
        if (auth.role !== 'supervisor') { result = { error: 'Supervisor access required' }; break; }
        result = approveSession(body.sessionId, body.approvedBy);
        break;
      case 'deleteSession':
        if (auth.role !== 'supervisor') { result = { error: 'Supervisor access required' }; break; }
        result = deleteSession(body.sessionId);
        break;
      case 'emailSessionPdf':
        if (auth.role === 'none') { result = { error: 'Unauthorised' }; break; }
        result = emailSessionPdf(body.sessionId, body.pdfBase64, body.filename, auth.email);
        break;
      case 'importBlockPlacement':
        if (auth.role !== 'supervisor') { result = { error: 'Supervisor access required' }; break; }
        result = importBlockPlacement(body.studentId, body.clinicName, body.fileBase64, body.filename);
        break;
      default:
        result = { error: 'Unknown action: ' + body.action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Sheets treats a cell value starting with +, -, =, or @ as a formula (the
 * classic Lotus-123 "+" convention Google Sheets still honours), even when
 * written via Apps Script's setValues()/appendRow() — not just when typed in
 * the UI. Free-text supervisor feedback often starts with "+" as a bullet
 * marker, which produced #ERROR! cells. Prefixing with an apostrophe forces
 * Sheets to store it as plain text (the apostrophe itself is not stored).
 */
function sanitizeSheetText(v) {
  var s = (v == null) ? '' : String(v);
  if (/^[+\-=@]/.test(s)) return "'" + s;
  return s;
}

/**
 * Looks for an existing Sessions row for the same student + same ID prefix
 * (SES- or STU-, so a supervisor's and student's entries for the same session
 * never collide with each other) submitted within the last few minutes with
 * matching date and identical hour values. Used to catch accidental
 * double-submits (double-click, retry after a slow response, resubmitting a
 * cached draft) without blocking two genuinely different sessions that happen
 * to fall on the same calendar day.
 * Returns the existing sessionId if a likely duplicate is found, else null.
 */
var DUPLICATE_SUBMISSION_WINDOW_MS = 3 * 60 * 1000; // 3 minutes

function findDuplicateSession(sessSheet, prefix, studentId, dateStr, hrsValues) {
  var data = sessSheet.getDataRange().getValues();
  var now = Date.now();
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var sid = String(row[1] || '');
    if (sid.indexOf(prefix) !== 0) continue;
    if (String(row[2] || '') !== String(studentId)) continue;
    var ts = row[0] instanceof Date ? row[0].getTime() : null;
    if (!ts || (now - ts) > DUPLICATE_SUBMISSION_WINDOW_MS) continue;
    if (String(row[3] || '') !== dateStr) continue;
    var match = true;
    for (var k = 0; k < hrsValues.length; k++) {
      if (Math.abs((Number(row[8 + k]) || 0) - (Number(hrsValues[k]) || 0)) > 0.01) { match = false; break; }
    }
    if (match) return sid;
  }
  return null;
}

/**
 * Writes one session submission to the Sessions and Ratings tabs.
 * Called from the supervisor feedback form.
 */
function submitSession(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessSheet  = ss.getSheetByName('Sessions');
  var ratsSheet  = ss.getSheetByName('Ratings');
  if (!sessSheet)  return { error: 'Sessions tab not found — check sheet setup.' };
  if (!ratsSheet)  return { error: 'Ratings tab not found — check sheet setup.' };

  var timestamp = new Date();
  var tz        = Session.getScriptTimeZone();
  var sessionId = 'SES-' + Utilities.formatDate(timestamp, tz, 'yyyyMMdd') +
                  '-' + Math.floor(Math.random() * 9000 + 1000);

  var hrs = d.hours || {};
  var hrsValues = [
    hrs['Adult Diagnostic-obs']          || 0,
    hrs['Adult Diagnostic-test']         || 0,
    hrs['Paediatric Diagnostic-obs']     || 0,
    hrs['Paediatric Diagnostic-test']    || 0,
    hrs['Adult Rehabilitation-obs']      || 0,
    hrs['Adult Rehabilitation-test']     || 0,
    hrs['Paediatric Rehabilitation-obs'] || 0,
    hrs['Paediatric Rehabilitation-test']|| 0,
    hrs['Other-obs']                     || 0,
    hrs['Other-test']                    || 0,
    hrs['ORL Observation']               || 0,
    hrs['SLT Observation']               || 0,
    hrs['Simulation']                    || 0,
    hrs['Clinical Supervision']          || 0
  ];
  var dateStr = d.date ? (d.date + (d.time ? ' ' + d.time : '')) : '';
  var dupId = findDuplicateSession(sessSheet, 'SES-', d.studentId, dateStr, hrsValues);
  if (dupId) {
    return { error: 'This looks like a duplicate — a session for this student on this date with the same hours was already submitted moments ago (ID: ' + dupId + '). Check the dashboard before resubmitting.' };
  }

  // One row in Sessions tab
  sessSheet.appendRow([
    timestamp,
    sessionId,
    d.studentId   || '',
    d.date ? (d.date + (d.time ? ' ' + d.time : '')) : '',
    d.sup1        || '',
    d.sup2        || '',
    d.location    || '',
    (d.activities || []).join(', '),
    hrs['Adult Diagnostic-obs']          || 0,
    hrs['Adult Diagnostic-test']         || 0,
    hrs['Paediatric Diagnostic-obs']     || 0,
    hrs['Paediatric Diagnostic-test']    || 0,
    hrs['Adult Rehabilitation-obs']      || 0,
    hrs['Adult Rehabilitation-test']     || 0,
    hrs['Paediatric Rehabilitation-obs'] || 0,
    hrs['Paediatric Rehabilitation-test']|| 0,
    hrs['Other-obs']                     || 0,
    hrs['Other-test']                    || 0,
    hrs['ORL Observation']               || 0,
    hrs['SLT Observation']               || 0,
    hrs['Simulation']                    || 0,
    hrs['Clinical Supervision']          || 0,
    sanitizeSheetText(d.fbWell),
    sanitizeSheetText(d.fbImprove),
    sanitizeSheetText(d.fbGeneral),
    ((d.subTypes || {})['Adult Diagnostic']          || []).join(', '),
    ((d.subTypes || {})['Paediatric Diagnostic']     || []).join(', '),
    ((d.subTypes || {})['Adult Rehabilitation']      || []).join(', '),
    ((d.subTypes || {})['Paediatric Rehabilitation'] || []).join(', '),
    ((d.subTypes || {})['Other']                     || []).join(', '),
    ((d.subTypes || {})['Simulation']                || []).join(', '),
    d.supervisorMnzas ? true : false                 // [31] MNZAS
  ]);

  // One row per rated skill in Ratings tab — collected and written in a single batch call
  var skills = d.skills || {};
  var ratingRows = [];
  Object.keys(skills).forEach(function (clinicType) {
    var ctSkills = skills[clinicType];
    Object.keys(ctSkills).forEach(function (skillId) {
      var s = ctSkills[skillId];
      if (!s.rating && !s.isPri && !s.isStr) return;
      ratingRows.push([
        timestamp,
        sessionId,
        d.studentId  || '',
        clinicType,
        d.milestone  || '',
        skillId,
        s.rating     || 0,
        s.isPri      ? true : false,
        s.isStr      ? true : false,
        sanitizeSheetText(s.comment)
      ]);
    });
  });
  if (ratingRows.length) {
    var startRow = ratsSheet.getLastRow() + 1;
    ratsSheet.getRange(startRow, 1, ratingRows.length, ratingRows[0].length).setValues(ratingRows);
  }

  // Optional: send notification email to student
  var studentsSheet = ss.getSheetByName('Students');
  if (studentsSheet && d.studentId) {
    var data = studentsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.studentId) && data[i][3]) {
        try {
          MailApp.sendEmail({
            to: data[i][3],
            subject: 'New clinical feedback — ' + (data[i][2] || d.studentId) + ' — ' + sessionId + ' — ' + (d.date || ''),
            body: 'Hi ' + data[i][2] + ',\n\nYour supervisor has submitted feedback for your session on ' +
                  (d.date || 'recent session') + '.\n\nLog in to the dashboard to view your progress.\n\nUoA Audiology'
          });
        } catch (mailErr) { /* email optional — don't fail the submission */ }
        break;
      }
    }
  }

  return { success: true, sessionId: sessionId };
}

/**
 * Student-submitted hours session (no skill ratings, approved=false until reflection submitted).
 */
function submitStudentHours(d, auth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessSheet = ss.getSheetByName('Sessions');
  if (!sessSheet) return { error: 'Sessions tab not found.' };

  var timestamp = new Date();
  var tz        = Session.getScriptTimeZone();
  var sessionId = 'STU-' + Utilities.formatDate(timestamp, tz, 'yyyyMMdd') +
                  '-' + Math.floor(Math.random() * 9000 + 1000);

  // Enforce student can only submit for themselves
  if (auth.studentId !== d.studentId) return { error: 'Unauthorised student ID' };

  var hrs = d.hours || {};
  var hrsValues = [
    hrs['Adult Diagnostic-obs']          || 0,
    hrs['Adult Diagnostic-test']         || 0,
    hrs['Paediatric Diagnostic-obs']     || 0,
    hrs['Paediatric Diagnostic-test']    || 0,
    hrs['Adult Rehabilitation-obs']      || 0,
    hrs['Adult Rehabilitation-test']     || 0,
    hrs['Paediatric Rehabilitation-obs'] || 0,
    hrs['Paediatric Rehabilitation-test']|| 0,
    hrs['Other-obs']                     || 0,
    hrs['Other-test']                    || 0,
    hrs['ORL Observation']               || 0,
    hrs['SLT Observation']               || 0,
    hrs['Simulation']                    || 0,
    hrs['Clinical Supervision']          || 0
  ];
  var dupId = findDuplicateSession(sessSheet, 'STU-', d.studentId, d.date || '', hrsValues);
  if (dupId) {
    return { error: 'This looks like a duplicate — hours for this date with the same values were already submitted moments ago (ID: ' + dupId + '). Check your session list before resubmitting.' };
  }

  sessSheet.appendRow([
    timestamp,
    sessionId,
    d.studentId   || '',
    d.date        || '',
    d.supervisorName || '',
    '',                                           // Supervisor2 — blank for student form
    d.location    || '',
    (d.activities || []).join(', '),
    hrs['Adult Diagnostic-obs']          || 0,
    hrs['Adult Diagnostic-test']         || 0,
    hrs['Paediatric Diagnostic-obs']     || 0,
    hrs['Paediatric Diagnostic-test']    || 0,
    hrs['Adult Rehabilitation-obs']      || 0,
    hrs['Adult Rehabilitation-test']     || 0,
    hrs['Paediatric Rehabilitation-obs'] || 0,
    hrs['Paediatric Rehabilitation-test']|| 0,
    hrs['Other-obs']                     || 0,
    hrs['Other-test']                    || 0,
    hrs['ORL Observation']               || 0,
    hrs['SLT Observation']               || 0,
    hrs['Simulation']                    || 0,
    hrs['Clinical Supervision']          || 0,
    '',                                           // Feedback_Well
    '',                                           // Feedback_Improve
    '',                                           // Feedback_General
    ((d.subTypes || {})['Adult Diagnostic']          || []).join(', '),
    ((d.subTypes || {})['Paediatric Diagnostic']     || []).join(', '),
    ((d.subTypes || {})['Adult Rehabilitation']      || []).join(', '),
    ((d.subTypes || {})['Paediatric Rehabilitation'] || []).join(', '),
    ((d.subTypes || {})['Other']                     || []).join(', '),
    ((d.subTypes || {})['Simulation']                || []).join(', '),
    d.supervisorMnzas ? true : false,             // [31] MNZAS
    false,                                        // [32] Approved — requires reflection submission
    ''                                            // [33] ApprovedBy
  ]);

  return { success: true, sessionId: sessionId };
}

/**
 * Returns cohort list and skill definitions.
 * Reads from "Config" and "Skills" tabs.
 */
function getConfig() {
  return getCached('config_v1', function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Read cohorts from Config tab
    var configSheet = ss.getSheetByName('Config');
    var configData = configSheet.getDataRange().getValues();
    var cohorts = [];
    var tz = Session.getScriptTimeZone();
    function fmtDate(v){ return v instanceof Date ? Utilities.formatDate(v,tz,'yyyy-MM-dd') : String(v||''); }
    for (var i = 1; i < configData.length; i++) {
      if (configData[i][0]) {
        cohorts.push({
          year:   String(configData[i][0]),
          label:  configData[i][1] || 'MAud ' + configData[i][0],
          active: configData[i][2] === true || configData[i][2] === 'TRUE' || configData[i][2] === 'Yes',
          s1End:  fmtDate(configData[i][3]),
          s2End:  fmtDate(configData[i][4]),
          y2End:  fmtDate(configData[i][5])
        });
      }
    }

    // Read skill definitions from Skills tab
    var skillsSheet = ss.getSheetByName('Skills');
    var skillsData = skillsSheet.getDataRange().getValues();
    var skills = [];
    for (var i = 1; i < skillsData.length; i++) {
      if (skillsData[i][0]) {
        skills.push({
          id: skillsData[i][0],
          name: skillsData[i][1],
          objective: skillsData[i][2],
          scope: skillsData[i][3] || 'all',
          expS1: Number(skillsData[i][4]) || 1,
          expS2: Number(skillsData[i][5]) || 1,
          expY2: Number(skillsData[i][6]) || 1
        });
      }
    }

    return { cohorts: cohorts, skills: skills };
  });
}

/**
 * Returns student list for a given cohort.
 * Reads from "Students" tab, filtered by cohort column.
 */
function getStudents(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Students');
  var data = sheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(cohort)) {
      students.push({
        id: String(data[i][0]),
        cohort: String(data[i][1]),
        name: data[i][2],
        email: data[i][3] || ''
      });
    }
  }

  return { cohort: cohort, students: students };
}

/**
 * Returns all ratings for a specific student.
 * Reads from "Ratings" tab.
 */
function getRatings(studentId, cohort) {
  if (!studentId) return { error: 'student parameter required' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();

  // Build sessionId → date map from Sessions tab (col 1 = SessionID, col 3 = Date)
  var sessionDateMap = {};
  var sessSheet = ss.getSheetByName('Sessions');
  if (sessSheet) {
    var sessData = sessSheet.getDataRange().getValues();
    for (var j = 1; j < sessData.length; j++) {
      var sid = String(sessData[j][1]);
      var dval = sessData[j][3];
      var dstr = (dval instanceof Date)
        ? Utilities.formatDate(dval, tz, 'yyyy-MM-dd HH:mm')
        : String(dval || '');
      sessionDateMap[sid] = dstr;
    }
  }

  var sheet = ss.getSheetByName('Ratings');
  var data = sheet.getDataRange().getValues();
  var ratings = [];

  // Ratings schema: Timestamp|SessionID|StudentID|ClinicType|Milestone|SkillID|Rating|IsPriority|IsStrength|Comment
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(studentId)) {
      var sessionId = String(data[i][1]);
      ratings.push({
        timestamp:  data[i][0],
        sessionId:  sessionId,
        studentId:  String(data[i][2]),
        clinicType: data[i][3],
        milestone:  data[i][4],
        skillId:    data[i][5],
        rating:     Number(data[i][6]),
        isPriority: data[i][7] === true || data[i][7] === 'TRUE',
        isStrength: data[i][8] === true || data[i][8] === 'TRUE',
        comment:    data[i][9] || '',
        date:       sessionDateMap[sessionId] || ''
      });
    }
  }

  return { studentId: studentId, ratings: ratings };
}

/**
 * Combines config + students + cohort_overview into a single response so
 * the dashboard's initial load is one round trip instead of three serial ones.
 */
function getDashboardInit(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };
  var config = getConfig();
  var studentsResult = getStudents(cohort);
  var overview = getCohortOverview(cohort);
  return {
    cohorts: config.cohorts,
    skills: config.skills,
    students: studentsResult.students || [],
    overview: overview.overview || []
  };
}

/**
 * Combines config + students + cohort_overview for the first active cohort into a
 * single response, so the dashboard's very first load (cohort dropdown + initial
 * cohort data) is one round trip instead of two serial ones (config, then dashboard_init).
 */
function getFullInit() {
  var config = getConfig();
  var active = (config.cohorts || []).filter(function (c) { return c.active; });
  var cohort = active.length ? active[0].year : ((config.cohorts || []).length ? config.cohorts[0].year : null);
  if (!cohort) {
    return { cohorts: config.cohorts, skills: config.skills, cohort: null, students: [], overview: [] };
  }
  var studentsResult = getStudents(cohort);
  var overview = getCohortOverview(cohort);
  return {
    cohorts: config.cohorts,
    skills: config.skills,
    cohort: cohort,
    students: studentsResult.students || [],
    overview: overview.overview || []
  };
}

/**
 * Returns a summary overview for all students in a cohort.
 * For each student, returns latest rating per skill per term.
 */
function getCohortOverview(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get student list
  var studentsResult = getStudents(cohort);
  var studentIdSet = {};
  var studentMap = {};
  studentsResult.students.forEach(function(s) { studentIdSet[s.id] = true; studentMap[s.id] = s.name; });

  // Get all ratings
  var sheet = ss.getSheetByName('Ratings');
  var data = sheet.getDataRange().getValues();

  // Build summary: { studentId: { milestone: { skillId: { rating, isPriority, isStrength } } } }
  // Ratings schema: Timestamp|SessionID|StudentID|ClinicType|Milestone|SkillID|Rating|IsPriority|IsStrength|Comment
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][2]);
    if (!studentIdSet[sid]) continue;

    var milestone = data[i][4];
    var skillId   = data[i][5];
    var rating    = Number(data[i][6]);
    var timestamp = data[i][0];

    if (!summary[sid]) summary[sid] = {};
    if (!summary[sid][milestone]) summary[sid][milestone] = {};

    // Keep the most recent rating per skill per milestone
    var existing = summary[sid][milestone][skillId];
    if (!existing || timestamp > existing.timestamp) {
      summary[sid][milestone][skillId] = {
        rating: rating,
        isPriority: data[i][7] === true || data[i][7] === 'TRUE',
        isStrength: data[i][8] === true || data[i][8] === 'TRUE',
        timestamp: timestamp
      };
    }
  }

  // Format output
  var overview = [];
  for (var sid in summary) {
    overview.push({
      studentId: sid,
      studentName: studentMap[sid] || sid,
      milestones: summary[sid]
    });
  }

  return { cohort: cohort, overview: overview };
}

/**
 * Returns all sessions for a student with hours breakdown and approval status.
 * Sessions schema cols: [0]Timestamp [1]SessionID [2]StudentID [3]Date [4]Sup1 [5]Sup2
 *   [6]Location [7]Activities [8-21]Hours [22-24]Feedback [25-30]SubTypes [31]MNZAS [32]Approved [33]ApprovedBy
 */
function getHours(studentId) {
  if (!studentId) return { error: 'student parameter required' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  var data = sheet.getDataRange().getValues();
  var sessions = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== String(studentId)) continue;
    sessions.push({
      sessionId:   String(data[i][1]),
      date:        (data[i][3] instanceof Date)
                     ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
                     : String(data[i][3] || ''),
      sup1:        data[i][4],
      sup2:        data[i][5],
      location:    data[i][6],
      activities:  data[i][7],
      hrs: {
        adultDxObs:    Number(data[i][8])  || 0,
        adultDxTest:   Number(data[i][9])  || 0,
        paedDxObs:     Number(data[i][10]) || 0,
        paedDxTest:    Number(data[i][11]) || 0,
        adultRehabObs: Number(data[i][12]) || 0,
        adultRehabTest:Number(data[i][13]) || 0,
        paedRehabObs:  Number(data[i][14]) || 0,
        paedRehabTest: Number(data[i][15]) || 0,
        otherObs:      Number(data[i][16]) || 0,
        otherTest:     Number(data[i][17]) || 0,
        orl:           Number(data[i][18]) || 0,
        slt:           Number(data[i][19]) || 0,
        simulation:    Number(data[i][20]) || 0,
        supervision:   Number(data[i][21]) || 0
      },
      mnzas:       data[i][31] === true || data[i][31] === 'TRUE',
      approved:    data[i][32] === true || data[i][32] === 'TRUE',
      approvedBy:  data[i][33] || ''
    });
  }
  return { studentId: studentId, sessions: sessions };
}

var HOURS_CATEGORY_KEYS = ['adultDxObs','adultDxTest','paedDxObs','paedDxTest',
  'adultRehabObs','adultRehabTest','paedRehabObs','paedRehabTest',
  'otherObs','otherTest','orl','slt','simulation','supervision'];

/**
 * Returns aggregated hours totals for all students in a cohort.
 * `totals` counts only approved sessions (these are what count toward compliance
 * targets); `pending` holds the same category sums for not-yet-approved sessions,
 * so the dashboard can show "X (+Y pending)" without those hours counting yet.
 */
function getCohortHours(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };
  var studentsResult = getStudents(cohort);
  var studentIds = studentsResult.students.map(function(s) { return s.id; });
  var studentMap = {};
  studentsResult.students.forEach(function(s) { studentMap[s.id] = s.name; });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  var data = sheet.getDataRange().getValues();

  function emptyCategoryTotals() {
    var o = {};
    HOURS_CATEGORY_KEYS.forEach(function(k) { o[k] = 0; });
    return o;
  }

  var totals = {}, pending = {};
  studentIds.forEach(function(id) {
    totals[id] = Object.assign(emptyCategoryTotals(), { approved: 0, sessions: 0 });
    pending[id] = emptyCategoryTotals();
  });

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][2]);
    if (!totals[sid]) continue;
    var isApproved = data[i][32] === true || data[i][32] === 'TRUE';
    totals[sid].sessions++;
    if (isApproved) totals[sid].approved++;
    var dest = isApproved ? totals[sid] : pending[sid];
    dest.adultDxObs    += Number(data[i][8])  || 0;
    dest.adultDxTest   += Number(data[i][9])  || 0;
    dest.paedDxObs     += Number(data[i][10]) || 0;
    dest.paedDxTest    += Number(data[i][11]) || 0;
    dest.adultRehabObs += Number(data[i][12]) || 0;
    dest.adultRehabTest+= Number(data[i][13]) || 0;
    dest.paedRehabObs  += Number(data[i][14]) || 0;
    dest.paedRehabTest += Number(data[i][15]) || 0;
    dest.otherObs       += Number(data[i][16]) || 0;
    dest.otherTest      += Number(data[i][17]) || 0;
    dest.orl             += Number(data[i][18]) || 0;
    dest.slt              += Number(data[i][19]) || 0;
    dest.simulation        += Number(data[i][20]) || 0;
    dest.supervision        += Number(data[i][21]) || 0;
  }

  var result = studentIds.map(function(id) {
    return { studentId: id, studentName: studentMap[id] || id, totals: totals[id], pending: pending[id] };
  });
  return { cohort: cohort, students: result };
}

/**
 * Returns every individual session (not aggregated) for all students in a cohort,
 * with studentName attached — powers the supervisor's "Cohort sessions" view so
 * they can scan/approve reflections across the whole cohort without drilling into
 * each student separately.
 */
function getCohortSessions(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };
  var studentsResult = getStudents(cohort);
  var studentMap = {};
  studentsResult.students.forEach(function (s) { studentMap[s.id] = s.name; });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var sessions = [];

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][2]);
    if (!studentMap.hasOwnProperty(sid)) continue;
    sessions.push({
      studentId: sid,
      studentName: studentMap[sid] || sid,
      sessionId: String(data[i][1]),
      date: (data[i][3] instanceof Date)
        ? Utilities.formatDate(data[i][3], tz, 'yyyy-MM-dd HH:mm')
        : String(data[i][3] || ''),
      sup1: data[i][4],
      sup2: data[i][5],
      location: data[i][6],
      activities: data[i][7],
      hrs: {
        adultDxObs:    Number(data[i][8])  || 0,
        adultDxTest:   Number(data[i][9])  || 0,
        paedDxObs:     Number(data[i][10]) || 0,
        paedDxTest:    Number(data[i][11]) || 0,
        adultRehabObs: Number(data[i][12]) || 0,
        adultRehabTest:Number(data[i][13]) || 0,
        paedRehabObs:  Number(data[i][14]) || 0,
        paedRehabTest: Number(data[i][15]) || 0,
        otherObs:      Number(data[i][16]) || 0,
        otherTest:     Number(data[i][17]) || 0,
        orl:           Number(data[i][18]) || 0,
        slt:           Number(data[i][19]) || 0,
        simulation:    Number(data[i][20]) || 0,
        supervision:   Number(data[i][21]) || 0
      },
      mnzas:       data[i][31] === true || data[i][31] === 'TRUE',
      approved:    data[i][32] === true || data[i][32] === 'TRUE',
      approvedBy:  data[i][33] || ''
    });
  }
  return { cohort: cohort, sessions: sessions };
}

/**
 * Returns one session's full record (all fields including feedback and
 * sub-types, which getHours/getCohortSessions don't read) plus every rating
 * row for that session (not filtered to flagged skills). Used by the
 * dashboard's session detail modal.
 * Sessions schema cols: [0]Timestamp [1]SessionID [2]StudentID [3]Date [4]Sup1 [5]Sup2
 *   [6]Location [7]Activities [8-21]Hours [22-24]Feedback [25-30]SubTypes [31]MNZAS [32]Approved [33]ApprovedBy
 */
function getSessionDetail(sessionId, auth) {
  if (!sessionId) return { error: 'sessionId parameter required' };
  if (!auth || auth.role === 'none') return { error: 'Unauthorised' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var sessSheet = ss.getSheetByName('Sessions');
  var data = sessSheet.getDataRange().getValues();
  var studentMap = {};
  var studentsSheet = ss.getSheetByName('Students');
  if (studentsSheet) {
    var sData = studentsSheet.getDataRange().getValues();
    for (var k = 1; k < sData.length; k++) studentMap[String(sData[k][0])] = sData[k][2];
  }

  var session = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) {
      var studentId = String(data[i][2]);
      session = {
        sessionId:   String(data[i][1]),
        studentId:   studentId,
        studentName: studentMap[studentId] || studentId,
        date:        (data[i][3] instanceof Date)
                       ? Utilities.formatDate(data[i][3], tz, 'yyyy-MM-dd HH:mm')
                       : String(data[i][3] || ''),
        sup1:        data[i][4],
        sup2:        data[i][5],
        location:    data[i][6],
        activities:  data[i][7],
        hrs: {
          adultDxObs:    Number(data[i][8])  || 0,
          adultDxTest:   Number(data[i][9])  || 0,
          paedDxObs:     Number(data[i][10]) || 0,
          paedDxTest:    Number(data[i][11]) || 0,
          adultRehabObs: Number(data[i][12]) || 0,
          adultRehabTest:Number(data[i][13]) || 0,
          paedRehabObs:  Number(data[i][14]) || 0,
          paedRehabTest: Number(data[i][15]) || 0,
          otherObs:      Number(data[i][16]) || 0,
          otherTest:     Number(data[i][17]) || 0,
          orl:           Number(data[i][18]) || 0,
          slt:           Number(data[i][19]) || 0,
          simulation:    Number(data[i][20]) || 0,
          supervision:   Number(data[i][21]) || 0
        },
        fbWell:      data[i][22] || '',
        fbImprove:   data[i][23] || '',
        fbGeneral:   data[i][24] || '',
        subTypes: {
          'Adult Diagnostic':          data[i][25] || '',
          'Paediatric Diagnostic':     data[i][26] || '',
          'Adult Rehabilitation':      data[i][27] || '',
          'Paediatric Rehabilitation': data[i][28] || '',
          'Other':                     data[i][29] || '',
          'Simulation':                data[i][30] || ''
        },
        mnzas:       data[i][31] === true || data[i][31] === 'TRUE',
        approved:    data[i][32] === true || data[i][32] === 'TRUE',
        approvedBy:  data[i][33] || ''
      };
      break;
    }
  }

  if (!session) return { error: 'Session not found' };
  if (auth.role === 'student' && auth.studentId !== session.studentId) return { error: 'Unauthorised' };

  var ratsSheet = ss.getSheetByName('Ratings');
  var ratings = [];
  if (ratsSheet) {
    var rData = ratsSheet.getDataRange().getValues();
    // Ratings schema: Timestamp|SessionID|StudentID|ClinicType|Milestone|SkillID|Rating|IsPriority|IsStrength|Comment
    for (var j = 1; j < rData.length; j++) {
      if (String(rData[j][1]) !== String(sessionId)) continue;
      ratings.push({
        clinicType: rData[j][3],
        milestone:  rData[j][4],
        skillId:    rData[j][5],
        rating:     Number(rData[j][6]),
        isPriority: rData[j][7] === true || rData[j][7] === 'TRUE',
        isStrength: rData[j][8] === true || rData[j][8] === 'TRUE',
        comment:    rData[j][9] || ''
      });
    }
  }

  return { session: session, ratings: ratings };
}

/**
 * Marks a session as approved. Writes TRUE + approver name to cols 26-27.
 */
function approveSession(sessionId, approvedBy) {
  if (!sessionId) return { error: 'sessionId required' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) {
      sheet.getRange(i + 1, 33).setValue(true);
      sheet.getRange(i + 1, 34).setValue(approvedBy || '');
      return { success: true, sessionId: sessionId };
    }
  }
  return { error: 'Session not found: ' + sessionId };
}

/**
 * Permanently deletes a session row (accidental/duplicate submission) plus any
 * Ratings rows tied to it. The client is required to double-confirm before
 * calling this — there is no undo once the rows are removed.
 */
function deleteSession(sessionId) {
  if (!sessionId) return { error: 'sessionId required' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sessions');
  if (!sheet) return { error: 'Sessions tab not found.' };

  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { error: 'Session not found: ' + sessionId };
  sheet.deleteRow(rowIndex);

  var ratsSheet = ss.getSheetByName('Ratings');
  if (ratsSheet) {
    var ratsData = ratsSheet.getDataRange().getValues();
    // Delete bottom-up so row indices of rows still to be checked aren't shifted.
    for (var j = ratsData.length - 1; j >= 1; j--) {
      if (String(ratsData[j][1]) === String(sessionId)) ratsSheet.deleteRow(j + 1);
    }
  }

  return { success: true, sessionId: sessionId };
}

/**
 * Emails the client-generated session PDF (base64) to the student and to
 * the relevant supervisors. Looks up the session row by sessionId so the
 * recipient list can't be spoofed by the client.
 * - SES- (supervisor feedback form): student + the signed-in submitting supervisor's
 *   preferred email (resolved from the Supervisors tab via their verified sign-in
 *   email — the free-text sup1/sup2 name fields can't be resolved to an email since
 *   the Supervisors tab has no Name column) + all coordinators
 * - STU- (student hours form):       student + all coordinators
 */
function emailSessionPdf(sessionId, pdfBase64, filename, submitterEmail) {
  if (!sessionId) return { error: 'sessionId required' };
  if (!pdfBase64) return { error: 'pdfBase64 required' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessSheet = ss.getSheetByName('Sessions');
  if (!sessSheet) return { error: 'Sessions tab not found.' };

  var data = sessSheet.getDataRange().getValues();
  var row = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) { row = data[i]; break; }
  }
  if (!row) return { error: 'Session not found: ' + sessionId };

  var studentId = String(row[2] || '');

  var studentsSheet = ss.getSheetByName('Students');
  var studentEmail = null, studentName = studentId;
  if (studentsSheet) {
    var stuData = studentsSheet.getDataRange().getValues();
    for (var j = 1; j < stuData.length; j++) {
      if (String(stuData[j][0]) === studentId) {
        studentEmail = stuData[j][3] || null;
        studentName = stuData[j][2] || studentId;
        break;
      }
    }
  }

  var dir = getSupervisorDirectory();
  var recipients = {};
  if (studentEmail) recipients[studentEmail.toLowerCase()] = true;
  dir.coordinatorEmails.forEach(function (e) { recipients[e.toLowerCase()] = true; });

  if (sessionId.indexOf('SES-') === 0 && submitterEmail) {
    var preferredEmail = dir.byLoginEmail[submitterEmail.toLowerCase()] || submitterEmail;
    recipients[preferredEmail.toLowerCase()] = true;
  }

  var emails = Object.keys(recipients);
  if (!emails.length) return { error: 'No recipients found for session ' + sessionId };

  var blob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', filename || (sessionId + '.pdf'));
  MailApp.sendEmail({
    to: emails.join(','),
    subject: 'MAud clinical session record — ' + studentName + ' — ' + sessionId + ' — ' + (row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), 'dd MMM yyyy') : ''),
    body: 'Attached is the session record PDF for session ' + sessionId + '.\n\nUoA Audiology',
    attachments: [blob]
  });

  return { success: true, sessionId: sessionId, recipients: emails };
}

/**
 * Utility: Add a rating row manually (bypasses the form submission flow).
 */
function addRating(sessionId, studentId, clinicType, milestone, skillId, rating, isPriority, isStrength, comment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ratings');
  sheet.appendRow([
    new Date(),
    sessionId,
    studentId,
    clinicType,
    milestone,
    skillId,
    rating,
    isPriority || false,
    isStrength || false,
    comment || ''
  ]);
}
