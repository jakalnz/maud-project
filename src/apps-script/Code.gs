/**
 * UoA Audiology Clinical Skills — Google Apps Script API
 *
 * SETUP:
 * 1. Create a Google Sheet with tabs: Config, Students, Skills, Sessions, Ratings
 *    (see docs/SHEET_SETUP.md for full column definitions)
 * 2. Open Extensions > Apps Script, paste this file into Code.gs
 * 3. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone within University of Auckland (or Anyone for external access)
 * 4. Copy the deployment URL and paste it into:
 *    - src/dashboard/index.html  →  const API_URL = '...';
 *    - src/form/index.html       →  const API_URL = '...';
 *
 * GET  ?action=config                           — cohorts + skill definitions
 * GET  ?action=students&cohort=2025             — student list
 * GET  ?action=ratings&student=STU001           — all ratings for a student
 * GET  ?action=cohort_overview&cohort=2025      — summary for all students
 * POST body: { action:'submitSession', data:{…} } — write session + ratings
 */

function doGet(e) {
  var action = e.parameter.action || 'config';
  var result;

  try {
    switch (action) {
      case 'config':
        result = getConfig();
        break;
      case 'students':
        result = getStudents(e.parameter.cohort);
        break;
      case 'ratings':
        result = getRatings(e.parameter.student, e.parameter.cohort);
        break;
      case 'cohort_overview':
        result = getCohortOverview(e.parameter.cohort);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests from the feedback form.
 * Content-Type must be text/plain (avoids CORS preflight).
 * Body: JSON string { action: 'submitSession', data: { … } }
 */
function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'submitSession':
        result = submitSession(body.data);
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
    d.date        || '',
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
    d.fbGeneral || ''
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
  var sheet = ss.getSheetByName('Ratings');
  var data = sheet.getDataRange().getValues();
  var ratings = [];

  // Ratings schema: Timestamp|SessionID|StudentID|ClinicType|Milestone|SkillID|Rating|IsPriority|IsStrength|Comment
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(studentId)) {
      ratings.push({
        timestamp:  data[i][0],
        sessionId:  String(data[i][1]),
        studentId:  String(data[i][2]),
        clinicType: data[i][3],
        milestone:  data[i][4],
        skillId:    data[i][5],
        rating:     Number(data[i][6]),
        isPriority: data[i][7] === true || data[i][7] === 'TRUE',
        isStrength: data[i][8] === true || data[i][8] === 'TRUE',
        comment:    data[i][9] || ''
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
