/**
 * UoA Audiology Clinical Skills Dashboard — Google Apps Script API
 * 
 * SETUP:
 * 1. Create a Google Sheet with the tabs described in SHEET_SETUP.md
 * 2. Open Extensions > Apps Script
 * 3. Paste this code into Code.gs
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (or "Anyone within [your org]" for university-only)
 * 5. Copy the deployment URL and paste it into dashboard.html where indicated
 * 
 * The script exposes a single doGet() endpoint that returns JSON.
 * Query params:
 *   ?action=config        — returns cohorts and skill definitions
 *   ?action=students&cohort=2025  — returns student list for a cohort
 *   ?action=ratings&student=STU001&cohort=2025  — returns all ratings for a student
 *   ?action=cohort_overview&cohort=2025  — returns summary for all students in a cohort
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
        clinicTypes: skillsData[i][3] ? skillsData[i][3].split(',').map(function(s) { return s.trim(); }) : ['All'],
        expT1: Number(skillsData[i][4]) || 1,
        expT2: Number(skillsData[i][5]) || 1,
        expT3: Number(skillsData[i][6]) || 1,
        expT4: Number(skillsData[i][7]) || 1
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

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(studentId)) {
      ratings.push({
        timestamp: data[i][0],
        studentId: String(data[i][1]),
        supervisor: data[i][2],
        clinicType: data[i][3],
        term: Number(data[i][4]),
        skillId: data[i][5],
        rating: Number(data[i][6]),
        isPriority: data[i][7] === true || data[i][7] === 'TRUE',
        isStrength: data[i][8] === true || data[i][8] === 'TRUE',
        comment: data[i][9] || ''
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

  // Build summary: { studentId: { term: { skillId: { rating, isPriority, isStrength } } } }
  var summary = {};

  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][1]);
    if (studentIds.indexOf(sid) === -1) continue;

    var term = Number(data[i][4]);
    var skillId = data[i][5];
    var rating = Number(data[i][6]);
    var timestamp = data[i][0];

    if (!summary[sid]) summary[sid] = {};
    if (!summary[sid][term]) summary[sid][term] = {};

    // Keep the most recent rating per skill per term
    var existing = summary[sid][term][skillId];
    if (!existing || timestamp > existing.timestamp) {
      summary[sid][term][skillId] = {
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
      terms: summary[sid]
    });
  }

  return { cohort: cohort, overview: overview };
}

/**
 * Utility: Add a rating row (can be called from a form or other script).
 */
function addRating(studentId, supervisor, clinicType, term, skillId, rating, isPriority, isStrength, comment) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Ratings');
  sheet.appendRow([
    new Date(),
    studentId,
    supervisor,
    clinicType,
    term,
    skillId,
    rating,
    isPriority || false,
    isStrength || false,
    comment || ''
  ]);
}
