/**
 * Cohort Hours Importer — one-off migration script
 *
 * Imports historical hours from "0. 2026 MAud Year 1 Clinical Hours.xlsx"
 * (one tab per student, dates across columns, activities down rows) directly
 * into the live Sessions tab used by the dashboard/Code.gs API.
 *
 * Unlike PlacementImporter.gs (which writes to a separate, unused "Hours"
 * tab), this writes Sessions rows in exactly the same column layout
 * submitSession() uses, so imported sessions show up in the dashboard like
 * any other session — with column 34 (AH) marked "*IMPORTED*" so they're
 * easy to find/filter later, and so re-running the import is safe (already
 * imported student+date pairs are skipped).
 *
 * USAGE (one-off, run manually from the Apps Script editor):
 * 1. Upload the xlsx to Google Drive, copy its file ID from the URL
 *    (the long string between /d/ and /edit).
 * 2. In the Apps Script editor: Services (+) > add "Drive API" (legacy v2,
 *    sometimes listed as "Drive API v2") as an Advanced Google Service.
 *    This file uses Drive.Files.insert to convert the uploaded xlsx to a
 *    temporary Google Sheet for parsing — without this service enabled you
 *    get "Drive is not defined".
 * 3. Edit the FILE_ID and COHORT_YEAR constants in runCohortHoursImportNow() below.
 * 4. Select "runCohortHoursImportNow" from the function dropdown and click Run.
 * 5. Check the alert dialog / Execution log for a summary (sessions created,
 *    unmatched student tabs, skipped dates).
 *
 * This file is self-contained — it does not depend on PlacementImporter.gs
 * (it has its own cohortImportOpenExcelAsSheet helper), so it works whether
 * or not that file is also pasted into the project.
 *
 * Student tabs are matched to the Students tab by name (the "X - " group
 * prefix in the tab name is stripped first). Tabs that don't match a
 * student are skipped and listed in the summary — add the student to the
 * Students tab (with their email) and re-run; already-imported dates are
 * skipped (no duplicate row created) automatically so re-running is safe.
 * For already-imported dates, the "Reflection Logged" value from the sheet
 * is re-checked against the existing row's Approved/ApprovedBy columns, and
 * the row is updated in place if it has changed (e.g. a clinician ticked
 * "yes" since the last import) — so re-uploading the same workbook after
 * clinicians log reflections keeps existing rows in sync rather than only
 * adding new ones.
 */

/**
 * Opens an Excel blob as a temporary Google Sheet for parsing.
 * Named distinctly from PlacementImporter.gs's openExcelAsSheet so this file
 * works standalone whether or not PlacementImporter.gs is also in the project.
 */
function cohortImportOpenExcelAsSheet(blob, name) {
  var resource = {
    title: '_temp_cohort_import_' + name,
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var file = Drive.Files.insert(resource, blob, { convert: true });
  return SpreadsheetApp.openById(file.id);
}

function runCohortHoursImportNow() {
  var FILE_ID = '1kKbRixKXir-CfHjHKWeCrfeMIQOdsZiX';
  var COHORT_YEAR = '2026';
  var result = importCohortHours(FILE_ID, COHORT_YEAR);
  SpreadsheetApp.getUi().alert(result.message);
  Logger.log(result.message);
}

// Tab names in the workbook that are not student tabs.
var COHORT_IMPORT_SKIP_TABS = ['GROUP INDEX', 'SUMMARY', 'Clinic Sign offs', 'Practicals', 'Reflections Tracker'];

// Excel row label -> { ct: clinic type, sub: SUBTYPES tag } for normal hours rows.
var COHORT_IMPORT_ACTIVITY_MAP = {
  'Full diagnostic hearing test (AC, BC, Speech, Immittance)': { ct: 'Adult Diagnostic', sub: 'Full diagnostic (AC/BC/Speech/Immittance)' },
  'Hearing screening (Pure-tone AC only)':                     { ct: 'Adult Diagnostic', sub: 'Hearing screening (AC only)' },
  'Extended High Frequency testing (Ototoxity, Tinnitus)':     { ct: 'Adult Diagnostic', sub: 'Extended HF / Tinnitus assessment' },
  'Tinnitus assessment (level and pitch matching)':            { ct: 'Adult Diagnostic', sub: 'Extended HF / Tinnitus assessment' },
  'APD testing (behavioural)':                                 { ct: 'Adult Diagnostic', sub: 'APD testing' },
  'Other objective tests (DPOAE, ABR, Corticals)':              { ct: 'Adult Diagnostic', sub: 'Other objective tests (DPOAE/ABR/Corticals)' },
  'Vestibular testing':                                        { ct: 'Adult Diagnostic', sub: 'Vestibular testing' },
  'Cochlear Implant assessment':                                { ct: 'Adult Diagnostic', sub: 'CI Assessment' },
  'Tele-audiology: diagnostics':                                { ct: 'Adult Diagnostic', sub: 'Tele-audiology' },

  'Needs Assessment':                                           { ct: 'Adult Rehabilitation', sub: 'Needs Assessment' },
  'Device Selection':                                           { ct: 'Adult Rehabilitation', sub: 'Device Selection' },
  'Hearing Aid Fitting / Real Ear Measures':                     { ct: 'Adult Rehabilitation', sub: 'HA Fitting / Real Ear Measures' },
  'Hearing aid follow up':                                       { ct: 'Adult Rehabilitation', sub: 'HA Follow-up' },
  'Tinnitus rehabilitation':                                     { ct: 'Adult Rehabilitation', sub: 'Tinnitus Rehabilitation' },
  'Cochlear Implant rehabilitation':                             { ct: 'Adult Rehabilitation', sub: 'CI Rehabilitation' },
  'Ear impressions, HA repairs/troubleshooting':                 { ct: 'Adult Rehabilitation', sub: 'Ear Impressions / HA Repairs' },
  'Tele-audiology: rehabilitation':                              { ct: 'Adult Rehabilitation', sub: 'Tele-audiology' },

  // Paediatric Diagnostic — note some labels collide with Adult Dx (e.g. "Cochlear Implant assessment").
  // Resolved by section instead: see parseCohortStudentSheet's currentSection tracking.
  'Newborn Hearing Screening (aABR)':                            { ct: 'Paediatric Diagnostic', sub: 'Newborn Hearing Screening (aABR)' },
  'Immittance (Tymp/ART) / DPOAE clinic':                        { ct: 'Paediatric Diagnostic', sub: 'Immittance / DPOAE clinic' },
  'Newborn Assessment (ABR)':                                    { ct: 'Paediatric Diagnostic', sub: 'Newborn Assessment (ABR)' },
  'Behavioural Observation Audiometry (BOA)':                    { ct: 'Paediatric Diagnostic', sub: 'BOA' },
  'Infant Assessment (VRA, MLV)':                                { ct: 'Paediatric Diagnostic', sub: 'Infant Assessment (VRA/MLV)' },
  'Preschool Assessments (CPA, KTT)':                            { ct: 'Paediatric Diagnostic', sub: 'Preschool Assessment (CPA/KTT)' },
  'School age child Assessments (PTA, CVC)':                     { ct: 'Paediatric Diagnostic', sub: 'School-age Assessment (PTA/CVC)' },
  'APD Testing (behavioural)':                                   { ct: 'Paediatric Diagnostic', sub: 'APD testing' },
  'Evoked Potentials':                                           { ct: 'Paediatric Diagnostic', sub: 'Evoked Potentials' },

  'Hearing Aid Fitting / Real Ear Measures / Test-box measures': { ct: 'Paediatric Rehabilitation', sub: 'HA Fitting / REM / Test-box' },

  'Observation of Ear Nurse Clinic':                             { ct: 'Other', sub: 'Observation – Ear Nurse' },
  'Observation of Hearing Therapist':                            { ct: 'Other', sub: 'Observation – Hearing Therapist / AODC / VHT' },
  'Observation of AODC/VHT':                                     { ct: 'Other', sub: 'Observation – Hearing Therapist / AODC / VHT' },
  'Equipment set-up and Troubleshooting':                        { ct: 'Other', sub: 'Equipment setup / Troubleshooting' },
  'Report writing and Case Discussions':                         { ct: 'Other', sub: 'Report writing / Case Discussion' },
  'HA Rep visit / seminar attendance/extra learning':            { ct: 'Other', sub: 'Seminar / HA Rep / Extra learning' }
};

// Labels that are ambiguous between Adult/Paediatric and Diagnostic/Rehab sections —
// resolved using the section the row falls under in the sheet, not a flat label lookup.
var COHORT_IMPORT_SECTION_MAP = {
  'Cochlear Implant assessment':       { sub: 'CI Assessment' },
  'Tele-audiology: diagnostics':       { sub: 'Tele-audiology' },
  'Cochlear Implant rehabilitation':   { sub: 'CI Rehabilitation' },
  'Ear impressions, HA repairs/troubleshooting': { sub: 'Ear Impressions / HA Repairs' },
  'Tele-audiology: rehabilitation':    { sub: 'Tele-audiology' },
  'Hearing aid follow up':             { sub: 'HA Follow-up' }
};

var COHORT_IMPORT_ORL_LABELS = ['Observation of ORL specialist in clinic', 'Observation of ORL surgery/theatre'];
var COHORT_IMPORT_SLT_LABELS = ['Observation of Speech Language Therapist'];
var COHORT_IMPORT_SUPERVISION_LABELS = ['Clinical Supervision of MAud Year 1 Students'];
var COHORT_IMPORT_SIMULATION_LABELS = ['Clinical Simulation'];

var COHORT_IMPORT_HOURS_COLS = {
  'Adult Diagnostic':         { obs: 'Adult Diagnostic-obs', test: 'Adult Diagnostic-test' },
  'Paediatric Diagnostic':    { obs: 'Paediatric Diagnostic-obs', test: 'Paediatric Diagnostic-test' },
  'Adult Rehabilitation':     { obs: 'Adult Rehabilitation-obs', test: 'Adult Rehabilitation-test' },
  'Paediatric Rehabilitation':{ obs: 'Paediatric Rehabilitation-obs', test: 'Paediatric Rehabilitation-test' },
  'Other':                    { obs: 'Other-obs', test: 'Other-test' }
};

function importCohortHours(fileId, cohortYear) {
  if (!fileId) return { message: 'fileId required.' };

  var file = DriveApp.getFileById(fileId);
  var tempBook = cohortImportOpenExcelAsSheet(file.getBlob(), file.getName());
  var sheets = tempBook.getSheets();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentsSheet = ss.getSheetByName('Students');
  var sessSheet = ss.getSheetByName('Sessions');
  if (!studentsSheet) return { message: 'Students tab not found.' };
  if (!sessSheet) return { message: 'Sessions tab not found.' };

  // Build name -> studentId map for the target cohort
  var stuData = studentsSheet.getDataRange().getValues();
  var nameToStudent = {};
  for (var i = 1; i < stuData.length; i++) {
    if (String(stuData[i][1]) === String(cohortYear)) {
      nameToStudent[String(stuData[i][2] || '').trim().toLowerCase()] = String(stuData[i][0]);
    }
  }

  // Build a map of "studentId|yyyy-MM-dd" -> sheet row (1-based) for already-imported
  // rows, so re-running is safe and can refresh the reflection/approval status.
  var existing = sessSheet.getDataRange().getValues();
  var alreadyImported = {};
  for (var r = 1; r < existing.length; r++) {
    if (existing[r][34] === '*IMPORTED*') {
      var d = existing[r][3];
      var dateKey = (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d).slice(0, 10);
      alreadyImported[existing[r][2] + '|' + dateKey] = { row: r + 1, approved: existing[r][32] === true };
    }
  }

  var sessionsCreated = 0, datesSkippedZero = 0, datesSkippedDNA = 0, datesSkippedDup = 0, reflectionsUpdated = 0;
  var unmatchedTabs = [];

  sheets.forEach(function (sheet) {
    var tabName = sheet.getName();
    if (COHORT_IMPORT_SKIP_TABS.indexOf(tabName) !== -1) return;

    var studentName = tabName.replace(/^[A-Za-z0-9]+\s*-\s*/, '').trim();
    var studentId = nameToStudent[studentName.toLowerCase()];
    if (!studentId) { unmatchedTabs.push(tabName); return; }

    var data = sheet.getDataRange().getValues();
    var parsed = parseCohortStudentSheet(data);

    parsed.dates.forEach(function (dateInfo) {
      var key = studentId + '|' + Utilities.formatDate(dateInfo.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var existingEntry = alreadyImported[key];
      if (existingEntry) {
        datesSkippedDup++;
        var nowApproved = dateInfo.reflectionLogged === true;
        if (nowApproved !== existingEntry.approved) {
          sessSheet.getRange(existingEntry.row, 33).setValue(nowApproved); // Approved
          sessSheet.getRange(existingEntry.row, 34).setValue(nowApproved ? (dateInfo.clinician || 'Imported') : ''); // ApprovedBy
          reflectionsUpdated++;
        }
        return;
      }
      if (dateInfo.didNotAttend) { datesSkippedDNA++; return; }

      var hrsTotal = 0;
      Object.keys(dateInfo.hours).forEach(function (k) { hrsTotal += dateInfo.hours[k] || 0; });
      hrsTotal += (dateInfo.orl || 0) + (dateInfo.slt || 0) + (dateInfo.supervision || 0) + (dateInfo.simulation || 0);
      if (hrsTotal <= 0) { datesSkippedZero++; return; }

      var activities = [];
      ['Adult Diagnostic', 'Paediatric Diagnostic', 'Adult Rehabilitation', 'Paediatric Rehabilitation', 'Other'].forEach(function (ct) {
        var cols = COHORT_IMPORT_HOURS_COLS[ct];
        if ((dateInfo.hours[cols.obs] || 0) + (dateInfo.hours[cols.test] || 0) > 0) activities.push(ct);
      });
      if (dateInfo.orl > 0) activities.push('ORL Observation');
      if (dateInfo.slt > 0) activities.push('SLT Observation');
      if (dateInfo.simulation > 0) activities.push('Simulation');

      var approved = dateInfo.reflectionLogged === true;
      sessSheet.appendRow([
        new Date(),
        'IMP-' + Utilities.formatDate(dateInfo.date, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + studentId.replace(/[^0-9A-Za-z]/g, ''),
        studentId,
        dateInfo.date,
        dateInfo.clinician || '',
        '',
        dateInfo.location || '',
        activities.join(', '),
        dateInfo.hours['Adult Diagnostic-obs'] || 0,
        dateInfo.hours['Adult Diagnostic-test'] || 0,
        dateInfo.hours['Paediatric Diagnostic-obs'] || 0,
        dateInfo.hours['Paediatric Diagnostic-test'] || 0,
        dateInfo.hours['Adult Rehabilitation-obs'] || 0,
        dateInfo.hours['Adult Rehabilitation-test'] || 0,
        dateInfo.hours['Paediatric Rehabilitation-obs'] || 0,
        dateInfo.hours['Paediatric Rehabilitation-test'] || 0,
        dateInfo.hours['Other-obs'] || 0,
        dateInfo.hours['Other-test'] || 0,
        dateInfo.orl || 0,
        dateInfo.slt || 0,
        dateInfo.simulation || 0,
        dateInfo.supervision || 0,
        '', '', '', // Feedback_Well / Improve / General — not present in the placement spreadsheet
        (dateInfo.subTypes['Adult Diagnostic'] || []).join(', '),
        (dateInfo.subTypes['Paediatric Diagnostic'] || []).join(', '),
        (dateInfo.subTypes['Adult Rehabilitation'] || []).join(', '),
        (dateInfo.subTypes['Paediatric Rehabilitation'] || []).join(', '),
        (dateInfo.subTypes['Other'] || []).join(', '),
        '', // SubTypes_Simulation — placement sheet doesn't break simulation into sub-types
        dateInfo.mnzas || false, // [31] MNZAS
        approved,                // [32] Approved
        approved ? (dateInfo.clinician || 'Imported') : '', // [33] ApprovedBy
        '*IMPORTED*'             // [34] ImportTag
      ]);
      sessionsCreated++;
    });
  });

  DriveApp.getFileById(tempBook.getId()).setTrashed(true);

  var message = 'Cohort hours import complete.\n\n' +
    'Sessions created: ' + sessionsCreated + '\n' +
    'Dates skipped (already imported): ' + datesSkippedDup + '\n' +
    '  ...of which reflection/approval status updated: ' + reflectionsUpdated + '\n' +
    'Dates skipped (Did Not Attend): ' + datesSkippedDNA + '\n' +
    'Dates skipped (no hours logged): ' + datesSkippedZero + '\n' +
    'Unmatched student tabs (add to Students tab + re-run): ' +
    (unmatchedTabs.length ? unmatchedTabs.join(', ') : 'none');

  return { message: message, sessionsCreated: sessionsCreated, reflectionsUpdated: reflectionsUpdated, unmatchedTabs: unmatchedTabs };
}

/**
 * Parses one student tab's full 2D values array into per-date aggregated hours.
 */
function parseCohortStudentSheet(data) {
  var result = { dates: [] };

  // Locate date columns from row 0: date at col c, 'O'/'T' at row1[c]/row1[c+1].
  var row0 = data[0], row1 = data[1] || [];
  var dateCols = []; // [{col, date}]
  for (var c = 2; c < row0.length; c++) {
    if (row0[c] instanceof Date) { dateCols.push({ col: c, date: row0[c] }); }
  }

  // Locate footer rows by label text (robust to minor row-number drift).
  var footerRow = { clinician: -1, mnzas: -1, location: -1, reflection: -1, dna: -1 };
  for (var r = 0; r < data.length; r++) {
    var label = String(data[r][0] || '').toLowerCase();
    if (label.indexOf('clinician full name') !== -1) footerRow.clinician = r;
    else if (label.indexOf('mnzas') !== -1) footerRow.mnzas = r;
    else if (label.indexOf('clinic name and location') !== -1) footerRow.location = r;
    else if (label.indexOf('reflection logged') !== -1) footerRow.reflection = r;
    else if (label.indexOf('did not attend') !== -1) footerRow.dna = r;
  }

  // Per-date aggregation buckets, keyed by column index.
  var byCol = {};
  dateCols.forEach(function (dc) {
    byCol[dc.col] = {
      date: dc.date, hours: {}, subTypes: { 'Adult Diagnostic': [], 'Paediatric Diagnostic': [], 'Adult Rehabilitation': [], 'Paediatric Rehabilitation': [], 'Other': [] },
      orl: 0, slt: 0, supervision: 0, simulation: 0,
      clinician: '', location: '', mnzas: false, reflectionLogged: false, didNotAttend: false
    };
  });

  // Track which top-level section (Adult/Paediatric) and sub-section (Diagnostic/Rehab) we're in,
  // to disambiguate row labels that repeat across sections (e.g. "Cochlear Implant assessment").
  var section = '', subSection = '';
  var footerRowsEnd = Math.min.apply(null, [footerRow.clinician, footerRow.location, footerRow.reflection, footerRow.dna].filter(function (v) { return v >= 0; }));

  for (var r = 2; r < (footerRowsEnd >= 0 ? footerRowsEnd : data.length); r++) {
    var label = String(data[r][0] || '').trim();
    if (!label) continue;
    if (label === 'ADULTS') { section = 'Adult'; continue; }
    if (label === 'PAEDIATRICS') { section = 'Paediatric'; continue; }
    if (label === 'OTHER') { section = 'Other'; continue; }
    if (label === 'Diagnostic') { subSection = 'Diagnostic'; continue; }
    if (label === 'Rehabilitation') { subSection = 'Rehabilitation'; continue; }
    if (label === 'Observations' || label === 'Other') continue; // sub-headers / subtotal rows

    var mapping = null, special = null;
    if (COHORT_IMPORT_ORL_LABELS.indexOf(label) !== -1) special = 'orl';
    else if (COHORT_IMPORT_SLT_LABELS.indexOf(label) !== -1) special = 'slt';
    else if (COHORT_IMPORT_SUPERVISION_LABELS.indexOf(label) !== -1) special = 'supervision';
    else if (COHORT_IMPORT_SIMULATION_LABELS.indexOf(label) !== -1) special = 'simulation';
    else if (COHORT_IMPORT_SECTION_MAP[label]) {
      var ct = (section === 'Paediatric' ? 'Paediatric' : 'Adult') + ' ' + (subSection === 'Rehabilitation' ? 'Rehabilitation' : 'Diagnostic');
      mapping = { ct: ct, sub: COHORT_IMPORT_SECTION_MAP[label].sub };
    } else if (COHORT_IMPORT_ACTIVITY_MAP[label]) {
      mapping = COHORT_IMPORT_ACTIVITY_MAP[label];
    } else {
      continue; // unrecognised row — skip rather than guess
    }

    dateCols.forEach(function (dc) {
      var bucket = byCol[dc.col];
      var obs = parseFloat(data[r][dc.col]) || 0;
      var test = parseFloat(data[r][dc.col + 1]) || 0;
      if (obs <= 0 && test <= 0) return;

      if (special === 'orl') { bucket.orl += obs + test; return; }
      if (special === 'slt') { bucket.slt += obs + test; return; }
      if (special === 'supervision') { bucket.supervision += obs + test; return; }
      if (special === 'simulation') { bucket.simulation += obs + test; return; }

      var cols = COHORT_IMPORT_HOURS_COLS[mapping.ct];
      bucket.hours[cols.obs] = (bucket.hours[cols.obs] || 0) + obs;
      bucket.hours[cols.test] = (bucket.hours[cols.test] || 0) + test;
      if (bucket.subTypes[mapping.ct].indexOf(mapping.sub) === -1) bucket.subTypes[mapping.ct].push(mapping.sub);
    });
  }

  // Fill in clinician/location/reflection/DNA per date column from footer rows.
  dateCols.forEach(function (dc) {
    var bucket = byCol[dc.col];
    if (footerRow.clinician >= 0) bucket.clinician = String(data[footerRow.clinician][dc.col] || '').trim();
    if (footerRow.mnzas >= 0) {
      var mnzasVal = String(data[footerRow.mnzas][dc.col] || '').trim().toLowerCase();
      bucket.mnzas = mnzasVal === 'y' || mnzasVal === 'yes' || mnzasVal === 'true';
    }
    if (footerRow.location >= 0) bucket.location = String(data[footerRow.location][dc.col] || '').trim();
    if (footerRow.reflection >= 0) bucket.reflectionLogged = String(data[footerRow.reflection][dc.col] || '').trim().toLowerCase() === 'yes';
    if (footerRow.dna >= 0) bucket.didNotAttend = !!String(data[footerRow.dna][dc.col] || '').trim();
    result.dates.push(bucket);
  });

  return result;
}
