/**
 * Block Placement Importer — Google Apps Script
 *
 * Handles the 'importBlockPlacement' POST action from the supervisor dashboard.
 * Supervisors upload the Block Placement Template Excel file directly from the
 * dashboard UI; the file arrives as a base64 string, is converted to a temporary
 * Google Sheet for parsing, and the resulting per-date sessions are written to
 * the live Sessions tab using the same column layout as submitSession() and
 * CohortHoursImporter.gs.
 *
 * Rows are tagged with '*BLOCK-PLACEMENT*' in column AH so they are easy to
 * identify and re-running is safe (existing studentId+date pairs from block
 * placements are skipped).
 *
 * Called from doPost() in Code.gs:
 *   POST { action:'importBlockPlacement', token, studentId, clinicName, fileBase64, filename }
 *   Returns { success, sessionsCreated, datesSkipped, totalHrs, message }
 */

// Activity label → { ct: clinic type, sub: sub-type label }
// Mirrors COHORT_IMPORT_ACTIVITY_MAP in CohortHoursImporter.gs but for the
// Block Placement Template row labels (same source doc, same labels).
var BPI_ACTIVITY_MAP = {
  // Adult Diagnostic
  'Full diagnostic hearing test (AC, BC, Speech, Immittance)': { ct: 'Adult Diagnostic', sub: 'Full diagnostic (AC/BC/Speech/Immittance)' },
  'Hearing screening (Pure-tone AC only)':                     { ct: 'Adult Diagnostic', sub: 'Hearing screening (AC only)' },
  'Extended High Frequency testing (Ototoxity, Tinnitus)':     { ct: 'Adult Diagnostic', sub: 'Extended HF / Tinnitus assessment' },
  'Tinnitus assessment (level and pitch matching)':            { ct: 'Adult Diagnostic', sub: 'Extended HF / Tinnitus assessment' },
  'APD testing (behavioural)':                                 { ct: 'Adult Diagnostic', sub: 'APD testing' },
  'Other objective tests (DPOAE, ABR, Corticals)':             { ct: 'Adult Diagnostic', sub: 'Other objective tests (DPOAE/ABR/Corticals)' },
  'Vestibular testing':                                        { ct: 'Adult Diagnostic', sub: 'Vestibular testing' },
  'Cochlear Implant assessment':                               { ct: 'Adult Diagnostic', sub: 'CI Assessment' },
  'Tele-audiology: diagnostics':                               { ct: 'Adult Diagnostic', sub: 'Tele-audiology' },

  // Adult Rehabilitation
  'Needs Assessment':                                          { ct: 'Adult Rehabilitation', sub: 'Needs Assessment' },
  'Device Selection':                                          { ct: 'Adult Rehabilitation', sub: 'Device Selection' },
  'Hearing Aid Fitting / Real Ear Measures':                   { ct: 'Adult Rehabilitation', sub: 'HA Fitting / Real Ear Measures' },
  'Hearing aid follow up':                                     { ct: 'Adult Rehabilitation', sub: 'HA Follow-up' },
  'Tinnitus rehabilitation':                                   { ct: 'Adult Rehabilitation', sub: 'Tinnitus Rehabilitation' },
  'Cochlear Implant rehabilitation':                           { ct: 'Adult Rehabilitation', sub: 'CI Rehabilitation' },
  'Ear impressions, HA repairs/troubleshooting':               { ct: 'Adult Rehabilitation', sub: 'Ear Impressions / HA Repairs' },
  'Tele-audiology: rehabilitation':                            { ct: 'Adult Rehabilitation', sub: 'Tele-audiology' },

  // Paediatric Diagnostic
  'Newborn Hearing Screening (aABR)':                          { ct: 'Paediatric Diagnostic', sub: 'Newborn Hearing Screening (aABR)' },
  'Immittance (Tymp/ART) / DPOAE clinic':                      { ct: 'Paediatric Diagnostic', sub: 'Immittance / DPOAE clinic' },
  'Newborn Assessment (ABR)':                                  { ct: 'Paediatric Diagnostic', sub: 'Newborn Assessment (ABR)' },
  'Behavioural Observation Audiometry (BOA)':                  { ct: 'Paediatric Diagnostic', sub: 'BOA' },
  'Infant Assessment (VRA, MLV)':                              { ct: 'Paediatric Diagnostic', sub: 'Infant Assessment (VRA/MLV)' },
  'Preschool Assessments (CPA, KTT)':                          { ct: 'Paediatric Diagnostic', sub: 'Preschool Assessment (CPA/KTT)' },
  'School age child Assessments (PTA, CVC)':                   { ct: 'Paediatric Diagnostic', sub: 'School-age Assessment (PTA/CVC)' },
  'APD Testing (behavioural)':                                 { ct: 'Paediatric Diagnostic', sub: 'APD testing' },
  'Evoked Potentials':                                         { ct: 'Paediatric Diagnostic', sub: 'Evoked Potentials' },

  // Paediatric Rehabilitation
  'Hearing Aid Fitting / Real Ear Measures / Test-box measures': { ct: 'Paediatric Rehabilitation', sub: 'HA Fitting / REM / Test-box' },

  // Other — observations go in here; ORL/SLT are handled separately below
  'Observation of Ear Nurse Clinic':                           { ct: 'Other', sub: 'Observation – Ear Nurse' },
  'Observation of Hearing Therapist':                          { ct: 'Other', sub: 'Observation – Hearing Therapist / AODC / VHT' },
  'Observation of AODC/VHT':                                   { ct: 'Other', sub: 'Observation – Hearing Therapist / AODC / VHT' },
  'Equipment set-up and Troubleshooting':                      { ct: 'Other', sub: 'Equipment setup / Troubleshooting' },
  'Report writing and Case Discussions':                       { ct: 'Other', sub: 'Report writing / Case Discussion' },
  'HA Rep visit / seminar attendance/extra learning':          { ct: 'Other', sub: 'Seminar / HA Rep / Extra learning' }
};

// Rows whose hours go into the ORL column (col S in Sessions, index 18)
var BPI_ORL_LABELS = [
  'Observation of ORL specialist in clinic',
  'Observation of ORL surgery/theatre'
];

// Rows whose hours go into the SLT column (col T, index 19)
var BPI_SLT_LABELS = [
  'Observation of Speech Language Therapist'
];

// Rows whose hours go into the Supervision column (col V, index 21)
var BPI_SUPERVISION_LABELS = [
  'Clinical Supervision of MAud Year 1 Students'
];

// Rows whose hours go into the Simulation column (col U, index 20)
var BPI_SIMULATION_LABELS = [
  'Clinical Simulation'
];

/**
 * Main entry point — called from doPost() in Code.gs.
 *
 * @param {string} studentId   - Target student ID (from Students tab)
 * @param {string} clinicName  - Placement clinic name (free text)
 * @param {string} fileBase64  - Base64-encoded Excel file content
 * @param {string} filename    - Original filename (used for temp sheet name)
 * @returns {{ success, sessionsCreated, datesSkipped, totalHrs, message }}
 */
function importBlockPlacement(studentId, clinicName, fileBase64, filename) {
  if (!studentId) return { success: false, message: 'studentId is required.' };
  if (!fileBase64) return { success: false, message: 'No file data received.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessSheet = ss.getSheetByName('Sessions');
  if (!sessSheet) return { success: false, message: 'Sessions tab not found — check sheet setup.' };

  // Decode base64 → blob → temporary Google Sheet
  var decoded = Utilities.base64Decode(fileBase64);
  var blob = Utilities.newBlob(decoded, MimeType.MICROSOFT_EXCEL, filename || 'placement.xlsx');
  var resource = { title: '_temp_bpi_' + (filename || 'placement'), mimeType: MimeType.GOOGLE_SHEETS };
  var tempFile = Drive.Files.insert(resource, blob, { convert: true });
  var tempBook = SpreadsheetApp.openById(tempFile.id);

  try {
    var hoursSheet = bpiFindClinicalHoursSheet(tempBook);
    if (!hoursSheet) {
      return { success: false, message: 'Could not find a "Clinical Hours" sheet in the uploaded file.' };
    }

    var data = hoursSheet.getDataRange().getValues();
    var parsed = bpiParseClinicalHours(data);

    if (parsed.dates.length === 0) {
      return { success: false, message: 'No date columns found in Clinical Hours sheet.' };
    }

    // Build set of already-imported block-placement dates for this student
    var existing = sessSheet.getDataRange().getValues();
    var alreadyImported = {};
    for (var r = 1; r < existing.length; r++) {
      if (existing[r][34] === '*BLOCK-PLACEMENT*' && String(existing[r][2]) === String(studentId)) {
        var d = existing[r][3];
        var dk = (d instanceof Date)
          ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(d).slice(0, 10);
        alreadyImported[dk] = true;
      }
    }

    var tz = Session.getScriptTimeZone();
    var sessionsCreated = 0, datesSkipped = 0, totalHrs = 0;

    parsed.dates.forEach(function (di) {
      var dateKey = Utilities.formatDate(di.date, tz, 'yyyy-MM-dd');
      if (alreadyImported[dateKey]) { datesSkipped++; return; }

      var hrs = di.hours;
      var rowTotal = 0;
      ['Adult Diagnostic-obs','Adult Diagnostic-test',
       'Paediatric Diagnostic-obs','Paediatric Diagnostic-test',
       'Adult Rehabilitation-obs','Adult Rehabilitation-test',
       'Paediatric Rehabilitation-obs','Paediatric Rehabilitation-test',
       'Other-obs','Other-test'].forEach(function (k) { rowTotal += hrs[k] || 0; });
      rowTotal += (di.orl || 0) + (di.slt || 0) + (di.supervision || 0) + (di.simulation || 0);
      if (rowTotal <= 0) { datesSkipped++; return; }

      var activities = [];
      ['Adult Diagnostic','Paediatric Diagnostic','Adult Rehabilitation','Paediatric Rehabilitation','Other'].forEach(function (ct) {
        if ((hrs[ct + '-obs'] || 0) + (hrs[ct + '-test'] || 0) > 0) activities.push(ct);
      });
      if (di.orl > 0)        activities.push('ORL Observation');
      if (di.slt > 0)        activities.push('SLT Observation');
      if (di.simulation > 0) activities.push('Simulation');

      var sessionId = 'BLK-' + Utilities.formatDate(di.date, tz, 'yyyyMMdd') + '-' + studentId.replace(/[^0-9A-Za-z]/g, '');

      sessSheet.appendRow([
        new Date(),                                          // col A  Timestamp
        sessionId,                                           // col B  SessionID
        studentId,                                           // col C  StudentID
        di.date,                                             // col D  Date
        di.clinician || clinicName || '',                    // col E  Sup1
        '',                                                  // col F  Sup2
        di.location || clinicName || '',                     // col G  Location
        activities.join(', '),                               // col H  Activities
        hrs['Adult Diagnostic-obs']          || 0,           // col I
        hrs['Adult Diagnostic-test']         || 0,           // col J
        hrs['Paediatric Diagnostic-obs']     || 0,           // col K
        hrs['Paediatric Diagnostic-test']    || 0,           // col L
        hrs['Adult Rehabilitation-obs']      || 0,           // col M
        hrs['Adult Rehabilitation-test']     || 0,           // col N
        hrs['Paediatric Rehabilitation-obs'] || 0,           // col O
        hrs['Paediatric Rehabilitation-test']|| 0,           // col P
        hrs['Other-obs']                     || 0,           // col Q
        hrs['Other-test']                    || 0,           // col R
        di.orl        || 0,                                  // col S  ORL
        di.slt        || 0,                                  // col T  SLT
        di.simulation || 0,                                  // col U  Simulation
        di.supervision|| 0,                                  // col V  Supervision
        '',                                                  // col W  FeedbackWell
        '',                                                  // col X  FeedbackImprove
        '',                                                  // col Y  FeedbackGeneral
        (di.subTypes['Adult Diagnostic']          || []).join(', '),  // col Z
        (di.subTypes['Paediatric Diagnostic']     || []).join(', '),  // col AA
        (di.subTypes['Adult Rehabilitation']      || []).join(', '),  // col AB
        (di.subTypes['Paediatric Rehabilitation'] || []).join(', '),  // col AC
        (di.subTypes['Other']                     || []).join(', '),  // col AD
        '',                                                  // col AE  SubTypes_Simulation
        di.mnzas || false,                                   // col AF  MNZAS
        false,                                               // col AG  Approved
        '',                                                  // col AH  ApprovedBy
        '*BLOCK-PLACEMENT*'                                  // col AI  ImportTag
      ]);

      sessionsCreated++;
      totalHrs += rowTotal;
    });

    return {
      success: true,
      sessionsCreated: sessionsCreated,
      datesSkipped: datesSkipped,
      totalHrs: Math.round(totalHrs * 10) / 10,
      message: 'Imported ' + sessionsCreated + ' session' + (sessionsCreated !== 1 ? 's' : '') +
               ' (' + datesSkipped + ' skipped — already imported or no hours).'
    };

  } finally {
    // Always clean up the temporary sheet
    DriveApp.getFileById(tempBook.getId()).setTrashed(true);
  }
}

/**
 * Finds the Clinical Hours sheet in the temporary workbook,
 * tolerating minor name variations.
 */
function bpiFindClinicalHoursSheet(book) {
  var names = ['Clinical Hours', 'Clinical hours', 'Hours', 'Sheet3'];
  for (var i = 0; i < names.length; i++) {
    var s = book.getSheetByName(names[i]);
    if (s) return s;
  }
  return null;
}

/**
 * Parses the 2D values array from the Clinical Hours sheet.
 *
 * Sheet layout (Block Placement Template):
 *   Row 0: "CLINICAL HOURS" | "Date of Clinic →" | <date> | <date> | ...
 *            (dates occupy one column each, but O/T pairs in row 1 tell us the
 *             observation column and test column for each date)
 *   Row 1: label | "Total Observed" | "Total Tested" | O | T | O | T | ...
 *   Rows 2–N: activity label | total-obs | total-test | obs | test | obs | test | ...
 *   Near bottom: "Clinician Full Name →" | | | name-per-date | ...
 *                "MNZAS? Y or N"         | ...
 *                "Clinic Name and Location →" | ...
 *
 * Returns { dates: [{ date, hours, subTypes, orl, slt, supervision, simulation, clinician, location }] }
 */
function bpiParseClinicalHours(data) {
  var result = { dates: [] };
  if (data.length < 2) return result;

  // --- Locate date columns ---
  // Row 0 has dates starting at col 2 (C); row 1 has O/T markers for each pair.
  var row0 = data[0];
  var row1 = data[1] || [];

  // Find O/T column pairs by scanning row 1 for the pattern "O" then "T".
  // The date for each pair is in row 0 at the same column as "O".
  var dateCols = []; // [{ obsCol, testCol, date }]
  for (var c = 2; c < row1.length - 1; c++) {
    var marker = String(row1[c] || '').trim().toUpperCase();
    var nextMarker = String(row1[c + 1] || '').trim().toUpperCase();
    if (marker === 'O' && nextMarker === 'T') {
      var rawDate = row0[c];
      // Some cells may have the date only in the first O column of the pair
      if (!rawDate && c > 2) rawDate = row0[c]; // already at obsCol
      if (rawDate instanceof Date) {
        dateCols.push({ obsCol: c, testCol: c + 1, date: rawDate });
      }
      c++; // advance past the T column
    }
  }

  if (dateCols.length === 0) return result;

  // Initialise per-date buckets
  var byCol = {};
  dateCols.forEach(function (dc) {
    byCol[dc.obsCol] = {
      date: dc.date,
      hours: {},
      subTypes: { 'Adult Diagnostic': [], 'Paediatric Diagnostic': [], 'Adult Rehabilitation': [], 'Paediatric Rehabilitation': [], 'Other': [] },
      orl: 0, slt: 0, supervision: 0, simulation: 0,
      clinician: '', location: '', mnzas: false
    };
  });

  // --- Scan activity rows ---
  // Track the current "section" for ambiguous labels (Cochlear Implant assessment
  // appears in both Adult Dx and Paed Dx sections — we resolve by current section).
  var currentSection = '';
  var SECTION_HEADERS = {
    'adults': null, 'adult': null,
    'paediatrics': null, 'paediatric': null,
    'other': null,
    'diagnostic': null,
    'rehabilitation': null, 'rehab': null
  };

  // Ambiguous labels resolved by currentSection
  var SECTION_RESOLVED = {
    'Cochlear Implant assessment':    { sub: 'CI Assessment' },
    'Tele-audiology: diagnostics':    { sub: 'Tele-audiology' },
    'Cochlear Implant rehabilitation':{ sub: 'CI Rehabilitation' },
    'Hearing aid follow up':          { sub: 'HA Follow-up' },
    'Ear impressions, HA repairs/troubleshooting': { sub: 'Ear Impressions / HA Repairs' },
    'Tele-audiology: rehabilitation': { sub: 'Tele-audiology' }
  };

  for (var r = 2; r < data.length; r++) {
    var label = String(data[r][0] || '').trim();
    var labelLower = label.toLowerCase();
    if (!label) continue;

    // --- Footer rows ---
    if (labelLower.indexOf('clinician full name') !== -1) {
      dateCols.forEach(function (dc) {
        var val = String(data[r][dc.obsCol] || '').trim();
        if (val) byCol[dc.obsCol].clinician = val;
      });
      continue;
    }
    if (labelLower.indexOf('mnzas') !== -1) {
      dateCols.forEach(function (dc) {
        var val = String(data[r][dc.obsCol] || '').trim().toLowerCase();
        byCol[dc.obsCol].mnzas = val === 'y' || val === 'yes' || val === 'true';
      });
      continue;
    }
    if (labelLower.indexOf('clinic name and location') !== -1) {
      dateCols.forEach(function (dc) {
        var val = String(data[r][dc.obsCol] || '').trim();
        if (val) byCol[dc.obsCol].location = val;
      });
      continue;
    }
    if (labelLower.indexOf('total hours') !== -1) continue;

    // --- Section headers ---
    if (labelLower === 'adults' || labelLower === 'adult') {
      currentSection = 'Adult';
      continue;
    }
    if (labelLower === 'paediatrics' || labelLower === 'paediatric') {
      currentSection = 'Paediatric';
      continue;
    }
    if (labelLower === 'other') {
      currentSection = 'Other';
      continue;
    }
    if (labelLower === 'diagnostic') {
      currentSection = (currentSection.split(' ')[0] || '') + ' Diagnostic';
      continue;
    }
    if (labelLower === 'rehabilitation' || labelLower === 'rehab') {
      currentSection = (currentSection.split(' ')[0] || '') + ' Rehabilitation';
      continue;
    }
    if (labelLower === 'observations') continue;

    // --- Special hour buckets ---
    var isOrl        = BPI_ORL_LABELS.indexOf(label) !== -1;
    var isSlt        = BPI_SLT_LABELS.indexOf(label) !== -1;
    var isSupervision= BPI_SUPERVISION_LABELS.indexOf(label) !== -1;
    var isSimulation = BPI_SIMULATION_LABELS.indexOf(label) !== -1;

    if (isOrl || isSlt || isSupervision || isSimulation) {
      dateCols.forEach(function (dc) {
        // These rows typically only use the obs column
        var hrs = parseFloat(data[r][dc.obsCol]) || 0;
        var bucket = byCol[dc.obsCol];
        if (isOrl)        bucket.orl        += hrs;
        if (isSlt)        bucket.slt        += hrs;
        if (isSupervision)bucket.supervision += hrs;
        if (isSimulation) bucket.simulation  += hrs;
      });
      continue;
    }

    // --- Regular activity rows ---
    var mapping = BPI_ACTIVITY_MAP[label];

    // If not in flat map, try section-resolved ambiguous labels
    if (!mapping && SECTION_RESOLVED[label] && currentSection) {
      var resolvedCt = currentSection; // e.g. "Adult Diagnostic"
      mapping = { ct: resolvedCt, sub: SECTION_RESOLVED[label].sub };
    }

    if (!mapping) continue; // Unknown row — skip

    var ct = mapping.ct;
    var sub = mapping.sub;
    var obsKey = ct + '-obs';
    var testKey = ct + '-test';

    dateCols.forEach(function (dc) {
      var obs  = parseFloat(data[r][dc.obsCol])  || 0;
      var test = parseFloat(data[r][dc.testCol]) || 0;
      if (obs <= 0 && test <= 0) return;

      var bucket = byCol[dc.obsCol];
      bucket.hours[obsKey]  = (bucket.hours[obsKey]  || 0) + obs;
      bucket.hours[testKey] = (bucket.hours[testKey] || 0) + test;
      if (sub && bucket.subTypes[ct] && bucket.subTypes[ct].indexOf(sub) === -1) {
        bucket.subTypes[ct].push(sub);
      }
    });
  }

  // Convert byCol map to ordered dates array
  dateCols.forEach(function (dc) {
    result.dates.push(byCol[dc.obsCol]);
  });

  return result;
}
