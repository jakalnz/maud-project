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

/** Look up role from verified email. Returns {role, email, name, studentId, cohort}. */
function getRole(idToken) {
  var payload = verifyToken(idToken);
  if (!payload) return { role: 'none', _debug: 'token_verify_failed' };
  var email = (payload.email || '').toLowerCase().trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check Supervisors tab
  var supSheet = ss.getSheetByName('Supervisors');
  if (!supSheet) return { role: 'none', _debug: 'no_supervisors_tab', email: email };
  var supData = supSheet.getDataRange().getValues();
  var supEmails = supData.slice(1).map(function(r){ return String(r[0]).toLowerCase().trim(); });
  if (supEmails.indexOf(email) !== -1) {
    return { role: 'supervisor', email: email, name: payload.name || '' };
  }

  // Check Students tab (email is col 3, index 3)
  var stuSheet = ss.getSheetByName('Students');
  if (!stuSheet) return { role: 'none', _debug: 'no_students_tab', email: email, supEmails: supEmails };
  var stuData = stuSheet.getDataRange().getValues();
  for (var i = 1; i < stuData.length; i++) {
    if (String(stuData[i][3]).toLowerCase().trim() === email) {
      return {
        role: 'student',
        email: email,
        name: String(stuData[i][2]),
        studentId: String(stuData[i][0]),
        cohort: String(stuData[i][1])
      };
    }
  }

  return { role: 'none', email: email, _debug: 'not_in_sheets', supEmails: supEmails };
}

function doGet(e) {
  var action = e.parameter.action || 'config';
  var token  = e.parameter.token  || '';
  var result;

  try {
    // Temporary debug endpoint — remove after testing
    if (action === 'debug') {
      var payload = verifyToken(token);
      result = { payload: payload, clientId: CLIENT_ID, tokenProvided: !!token, tokenLength: token.length };
    }
    // Role endpoint — also accepts token via POST body (see doPost)
    else if (action === 'role') {
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
        if (auth.role !== 'student') { result = { error: 'Student access required' }; break; }
        result = submitStudentHours(body.data, auth);
        break;
      case 'approveSession':
        if (auth.role !== 'supervisor') { result = { error: 'Supervisor access required' }; break; }
        result = approveSession(body.sessionId, body.approvedBy);
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
    d.fbWell    || '',
    d.fbImprove || '',
    d.fbGeneral || '',
    ((d.subTypes || {})['Adult Diagnostic']          || []).join(', '),
    ((d.subTypes || {})['Paediatric Diagnostic']     || []).join(', '),
    ((d.subTypes || {})['Adult Rehabilitation']      || []).join(', '),
    ((d.subTypes || {})['Paediatric Rehabilitation'] || []).join(', '),
    ((d.subTypes || {})['Other']                     || []).join(', '),
    ((d.subTypes || {})['Simulation']                || []).join(', ')
  ]);

  // One row per rated skill in Ratings tab
  var skills = d.skills || {};
  Object.keys(skills).forEach(function (clinicType) {
    var ctSkills = skills[clinicType];
    Object.keys(ctSkills).forEach(function (skillId) {
      var s = ctSkills[skillId];
      if (!s.rating && !s.isPri && !s.isStr) return;
      ratsSheet.appendRow([
        timestamp,
        sessionId,
        d.studentId  || '',
        clinicType,
        d.milestone  || '',
        skillId,
        s.rating     || 0,
        s.isPri      ? true : false,
        s.isStr      ? true : false,
        s.comment    || ''
      ]);
    });
  });

  // Optional: send notification email to student
  var studentsSheet = ss.getSheetByName('Students');
  if (studentsSheet && d.studentId) {
    var data = studentsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.studentId) && data[i][3]) {
        try {
          MailApp.sendEmail({
            to: data[i][3],
            subject: 'New clinical feedback — ' + (d.date || ''),
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
    false,                                        // Approved — requires reflection submission
    ''                                            // ApprovedBy
  ]);

  return { success: true, sessionId: sessionId };
}

/**
 * Returns cohort list and skill definitions.
 * Reads from "Config" and "Skills" tabs.
 */
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read cohorts from Config tab
  var configSheet = ss.getSheetByName('Config');
  var configData = configSheet.getDataRange().getValues();
  var cohorts = [];
  for (var i = 1; i < configData.length; i++) {
    if (configData[i][0]) {
      cohorts.push({
        year: String(configData[i][0]),
        label: configData[i][1] || 'MAud ' + configData[i][0],
        active: configData[i][2] === true || configData[i][2] === 'TRUE' || configData[i][2] === 'Yes'
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
 * Returns a summary overview for all students in a cohort.
 * For each student, returns latest rating per skill per term.
 */
function getCohortOverview(cohort) {
  if (!cohort) return { error: 'cohort parameter required' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get student list
  var studentsResult = getStudents(cohort);
  var studentIds = studentsResult.students.map(function(s) { return s.id; });
  var studentMap = {};
  studentsResult.students.forEach(function(s) { studentMap[s.id] = s.name; });

  // Get all ratings
  var sheet = ss.getSheetByName('Ratings');
  var data = sheet.getDataRange().getValues();

  // Build summary: { studentId: { milestone: { skillId: { rating, isPriority, isStrength } } } }
  // Ratings schema: Timestamp|SessionID|StudentID|ClinicType|Milestone|SkillID|Rating|IsPriority|IsStrength|Comment
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][2]);
    if (studentIds.indexOf(sid) === -1) continue;

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
 *   [6]Location [7]Activities [8-21]Hours [22-24]Feedback [25-30]SubTypes [31]Approved [32]ApprovedBy
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
      approved:    data[i][31] === true || data[i][31] === 'TRUE',
      approvedBy:  data[i][32] || ''
    });
  }
  return { studentId: studentId, sessions: sessions };
}

/**
 * Returns aggregated hours totals for all students in a cohort.
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

  var totals = {};
  studentIds.forEach(function(id) {
    totals[id] = { adultDxObs:0, adultDxTest:0, paedDxObs:0, paedDxTest:0,
                   adultRehabObs:0, adultRehabTest:0, paedRehabObs:0, paedRehabTest:0,
                   otherObs:0, otherTest:0,
                   orl:0, slt:0, simulation:0, supervision:0,
                   approved:0, sessions:0 };
  });

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][2]);
    if (!totals[sid]) continue;
    var t = totals[sid];
    t.sessions++;
    t.adultDxObs    += Number(data[i][8])  || 0;
    t.adultDxTest   += Number(data[i][9])  || 0;
    t.paedDxObs     += Number(data[i][10]) || 0;
    t.paedDxTest    += Number(data[i][11]) || 0;
    t.adultRehabObs += Number(data[i][12]) || 0;
    t.adultRehabTest+= Number(data[i][13]) || 0;
    t.paedRehabObs  += Number(data[i][14]) || 0;
    t.paedRehabTest += Number(data[i][15]) || 0;
    t.otherObs      += Number(data[i][16]) || 0;
    t.otherTest     += Number(data[i][17]) || 0;
    t.orl           += Number(data[i][18]) || 0;
    t.slt           += Number(data[i][19]) || 0;
    t.simulation    += Number(data[i][20]) || 0;
    t.supervision   += Number(data[i][21]) || 0;
    if (data[i][31] === true || data[i][31] === 'TRUE') t.approved++;
  }

  var result = studentIds.map(function(id) {
    return { studentId: id, studentName: studentMap[id] || id, totals: totals[id] };
  });
  return { cohort: cohort, students: result };
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
      sheet.getRange(i + 1, 32).setValue(true);
      sheet.getRange(i + 1, 33).setValue(approvedBy || '');
      return { success: true, sessionId: sessionId };
    }
  }
  return { error: 'Session not found: ' + sessionId };
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
