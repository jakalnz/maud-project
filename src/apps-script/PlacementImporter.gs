/**
 * Placement Hours Importer — Google Apps Script
 * 
 * Add this to your existing Apps Script project (alongside Code.gs).
 * 
 * USAGE:
 * 1. Upload the placement Excel file to Google Drive
 * 2. From the Google Sheet menu: Placements > Import Placement Hours
 * 3. Select the uploaded file from the picker
 * 4. The script parses the Clinical Hours sheet and writes to the Hours tab
 * 
 * Alternatively, call importPlacementFromDrive(fileId, studentId) programmatically.
 */

/**
 * Adds a custom menu to the spreadsheet for easy access.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Placements')
    .addItem('Import Placement Hours...', 'showImportDialog')
    .addItem('View Import Log', 'showImportLog')
    .addSeparator()
    .addItem('List Pending Placement Reviews', 'listPendingReviews')
    .addToUi();
}

/**
 * Shows a dialog for selecting the placement file and student.
 */
function showImportDialog() {
  var html = HtmlService.createHtmlOutput(getImportDialogHtml())
    .setWidth(450)
    .setHeight(400)
    .setTitle('Import Placement Hours');
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Placement Hours');
}

function getImportDialogHtml() {
  // Get student list for dropdown
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentsSheet = ss.getSheetByName('Students');
  var studentsData = studentsSheet.getDataRange().getValues();
  var studentOptions = '';
  for (var i = 1; i < studentsData.length; i++) {
    if (studentsData[i][0]) {
      studentOptions += '<option value="' + studentsData[i][0] + '">' + 
        studentsData[i][2] + ' (' + studentsData[i][1] + ')</option>';
    }
  }
  
  return '<html><head><style>' +
    'body{font-family:Arial,sans-serif;padding:16px}' +
    'label{display:block;font-size:13px;color:#555;margin:12px 0 4px}' +
    'select,input{width:100%;padding:8px;font-size:13px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box}' +
    '.btn{padding:10px 20px;font-size:14px;border:none;border-radius:6px;cursor:pointer;margin-top:16px}' +
    '.btn-primary{background:#185FA5;color:white}' +
    '.btn-primary:hover{background:#134d8a}' +
    '#status{margin-top:12px;padding:10px;border-radius:4px;font-size:13px;display:none}' +
    '.success{background:#EAF3DE;color:#3B6D11}' +
    '.error{background:#FCEBEB;color:#A32D2D}' +
    '.info{background:#E6F1FB;color:#0C447C}' +
    '</style></head><body>' +
    '<h3 style="margin:0 0 4px;color:#1A365D">Import Placement Hours</h3>' +
    '<p style="font-size:12px;color:#888;margin:0 0 16px">Select the student and paste the Google Drive file ID of their placement spreadsheet.</p>' +
    '<label>Student</label>' +
    '<select id="studentId">' + studentOptions + '</select>' +
    '<label>Placement clinic name</label>' +
    '<input type="text" id="clinicName" placeholder="e.g. Bay Audiology, Hamilton">' +
    '<label>Term</label>' +
    '<select id="term">' +
    '<option value="1">Term 1</option><option value="2">Term 2</option>' +
    '<option value="3">Term 3</option><option value="4">Term 4</option>' +
    '<option value="SB">Summer Block</option></select>' +
    '<label>Google Drive File ID</label>' +
    '<input type="text" id="fileId" placeholder="e.g. 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms">' +
    '<p style="font-size:11px;color:#999;margin:4px 0 0">The file ID is the long string in the Google Drive URL between /d/ and /edit</p>' +
    '<button class="btn btn-primary" onclick="doImport()">Import Hours</button>' +
    '<div id="status"></div>' +
    '<script>' +
    'function doImport(){' +
    '  var s=document.getElementById("status");' +
    '  s.style.display="block";s.className="info";s.innerHTML="Importing... please wait.";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      s.className=r.success?"success":"error";' +
    '      s.innerHTML=r.message;' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      s.className="error";s.innerHTML="Error: "+e.message;' +
    '    })' +
    '    .importPlacementFromDrive(' +
    '      document.getElementById("fileId").value,' +
    '      document.getElementById("studentId").value,' +
    '      document.getElementById("clinicName").value,' +
    '      document.getElementById("term").value' +
    '    );' +
    '}' +
    '</script></body></html>';
}

/**
 * Main import function. Reads the placement Excel file from Drive,
 * parses the Clinical Hours sheet, and writes to the Hours tab.
 * 
 * @param {string} fileId - Google Drive file ID of the placement Excel
 * @param {string} studentId - Student ID from the Students tab
 * @param {string} clinicName - Name of the placement clinic
 * @param {string} term - Term identifier (1-4 or SB for Summer Block)
 * @returns {object} - {success: boolean, message: string}
 */
function importPlacementFromDrive(fileId, studentId, clinicName, term) {
  try {
    // Validate inputs
    if (!fileId || !studentId) {
      return { success: false, message: 'File ID and Student are required.' };
    }
    
    // Open the uploaded Excel file as a temporary Google Sheet
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var tempSheet = openExcelAsSheet(blob, file.getName());
    
    // Parse the Clinical Hours sheet
    var hoursSheet = tempSheet.getSheetByName('Clinical Hours');
    if (!hoursSheet) {
      // Try common variations
      hoursSheet = tempSheet.getSheetByName('Clinical hours') || 
                   tempSheet.getSheetByName('Hours') ||
                   tempSheet.getSheetByName('Sheet3');
      if (!hoursSheet) {
        DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
        return { success: false, message: 'Could not find "Clinical Hours" sheet in the uploaded file.' };
      }
    }
    
    var data = hoursSheet.getDataRange().getValues();
    var parsed = parseClinicalHours(data);
    
    // Clean up temp file
    DriveApp.getFileById(tempSheet.getId()).setTrashed(true);
    
    if (parsed.entries.length === 0) {
      return { success: false, message: 'No hours data found in the file. Check that dates and hours are filled in.' };
    }
    
    // Write to the Hours tab
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoursTab = ss.getSheetByName('Hours');
    if (!hoursTab) {
      // Create Hours tab if it doesn't exist
      hoursTab = ss.insertSheet('Hours');
      hoursTab.appendRow([
        'Timestamp', 'StudentID', 'Supervisor', 'ClinicType', 'Term',
        'Date', 'Activity', 'ObservedHrs', 'TestedHrs',
        'Source', 'PlacementClinic', 'MNZAS', 'SessionID', 'ReflectionConfirmed'
      ]);
    }
    
    var rowsAdded = 0;
    var totalObserved = 0;
    var totalTested = 0;
    
    parsed.entries.forEach(function(entry) {
      // Determine clinic type from the activity category
      var clinicType = categoriseActivity(entry.category, entry.activity);
      
      // Generate a session ID for this date
      var sessionId = 'PLC-' + studentId.replace('STU', '') + '-' + 
        formatDateForId(entry.date);
      
      // Get supervisor for this date
      var supervisor = parsed.supervisors[entry.dateIndex] || 'Unknown';
      var mnzas = parsed.mnzas[entry.dateIndex] || '';
      
      hoursTab.appendRow([
        new Date(),           // Timestamp
        studentId,            // StudentID
        supervisor,           // Supervisor
        clinicType,           // ClinicType
        term,                 // Term
        entry.date,           // Date
        entry.activity,       // Activity
        entry.observed,       // ObservedHrs
        entry.tested,         // TestedHrs
        'Placement',          // Source
        clinicName,           // PlacementClinic
        mnzas,                // MNZAS
        sessionId,            // SessionID
        false                 // ReflectionConfirmed
      ]);
      
      rowsAdded++;
      totalObserved += entry.observed;
      totalTested += entry.tested;
    });
    
    // Log the import
    logImport(studentId, clinicName, term, rowsAdded, totalObserved, totalTested, fileId);
    
    return {
      success: true,
      message: 'Successfully imported ' + rowsAdded + ' hour entries.<br>' +
        'Total observed: ' + totalObserved.toFixed(1) + ' hrs<br>' +
        'Total tested: ' + totalTested.toFixed(1) + ' hrs<br>' +
        'Total: ' + (totalObserved + totalTested).toFixed(1) + ' hrs<br>' +
        'Dates: ' + parsed.dates.length + ' clinic days found.'
    };
    
  } catch (e) {
    return { success: false, message: 'Import failed: ' + e.message };
  }
}

/**
 * Opens an Excel blob as a temporary Google Sheet for parsing.
 */
function openExcelAsSheet(blob, name) {
  var resource = {
    title: '_temp_import_' + name,
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var file = Drive.Files.insert(resource, blob, { convert: true });
  return SpreadsheetApp.openById(file.id);
}

/**
 * Parses the Clinical Hours sheet data into structured entries.
 * 
 * The sheet structure is:
 * - Row 1: Header row with dates starting from column C (alternating O/T pairs)
 * - Row 2: "Total Observed", "Total Tested", then O, T, O, T... per date
 * - Rows 3+: Activity names in col A, with hours in the O/T columns
 * - Near-bottom rows: Clinician name, MNZAS status, Clinic location
 * - Last row: Total hours per day
 * 
 * @param {Array[]} data - 2D array of cell values
 * @returns {object} - {entries: [], dates: [], supervisors: {}, mnzas: {}}
 */
function parseClinicalHours(data) {
  var result = {
    entries: [],
    dates: [],
    supervisors: {},
    mnzas: {}
  };
  
  if (data.length < 3) return result;
  
  // Find date columns - dates are in the first row, starting from column index 3+
  // The structure has "Total Observed" and "Total Tested" in cols B/C,
  // then alternating O/T columns for each date
  var headerRow = data[0];
  var dateColumns = {}; // dateIndex -> {col: colIndex, date: dateValue}
  var dateIndex = 0;
  
  for (var c = 2; c < headerRow.length; c++) {
    var cellVal = headerRow[c];
    if (cellVal instanceof Date || isDateLike(cellVal)) {
      var dateVal = cellVal instanceof Date ? cellVal : new Date(cellVal);
      if (!isNaN(dateVal.getTime())) {
        dateColumns[dateIndex] = { obsCol: c, testCol: c + 1, date: dateVal };
        result.dates.push(dateVal);
        dateIndex++;
        c++; // Skip the T column (already captured as testCol)
      }
    }
  }
  
  // If no dates found in row 0, try row 1 (some templates have a different header layout)
  if (dateIndex === 0 && data.length > 1) {
    headerRow = data[1];
    // Look for O/T pattern to identify date column pairs
    for (var c = 2; c < headerRow.length; c++) {
      if (String(headerRow[c]).trim().toUpperCase() === 'O' && 
          c + 1 < headerRow.length && 
          String(headerRow[c + 1]).trim().toUpperCase() === 'T') {
        // Check the row above for a date
        var possibleDate = data[0][c];
        if (possibleDate instanceof Date || isDateLike(possibleDate)) {
          var dateVal = possibleDate instanceof Date ? possibleDate : new Date(possibleDate);
          if (!isNaN(dateVal.getTime())) {
            dateColumns[dateIndex] = { obsCol: c, testCol: c + 1, date: dateVal };
            result.dates.push(dateVal);
            dateIndex++;
          }
        }
        c++; // Skip T column
      }
    }
  }
  
  if (dateIndex === 0) return result; // No dates found
  
  // Define the activity categories and their hierarchy
  // We track the current parent category (Adults Diagnostic, Adults Rehab, etc.)
  var currentCategory = '';
  var supervisorRowLabel = ['clinician full name', 'clinician name', 'supervisor'];
  var mnzasRowLabel = ['mnzas', 'mnzas?'];
  var skipRows = ['adults', 'paediatrics', 'other', 'observations',
                  'diagnostic', 'rehabilitation', 'total hours'];
  
  // Parse activity rows (typically rows 2 onwards, after headers)
  for (var r = 2; r < data.length; r++) {
    var label = String(data[r][0] || '').trim();
    var labelLower = label.toLowerCase();
    
    // Check for supervisor/MNZAS info rows
    if (supervisorRowLabel.some(function(s) { return labelLower.indexOf(s) !== -1; })) {
      for (var di in dateColumns) {
        var dc = dateColumns[di];
        var supName = String(data[r][dc.obsCol] || '').trim();
        if (supName) result.supervisors[di] = supName;
      }
      continue;
    }
    
    if (mnzasRowLabel.some(function(s) { return labelLower.indexOf(s) !== -1; })) {
      for (var di in dateColumns) {
        var dc = dateColumns[di];
        var mnzas = String(data[r][dc.obsCol] || '').trim();
        if (mnzas) result.mnzas[di] = mnzas;
      }
      continue;
    }
    
    // Skip category headers and totals rows
    if (!label || skipRows.some(function(s) { return labelLower === s; })) {
      // But track category changes
      if (labelLower === 'diagnostic' || labelLower === 'rehabilitation') {
        // Use in combination with the parent (Adults/Paediatrics)
      }
      continue;
    }
    
    // Track parent categories for clinic type mapping
    if (labelLower === 'adults' || labelLower === 'adult') {
      currentCategory = 'Adult';
      continue;
    } else if (labelLower === 'paediatrics' || labelLower === 'paediatric') {
      currentCategory = 'Paediatric';
      continue;
    } else if (labelLower === 'other') {
      currentCategory = 'Other';
      continue;
    } else if (labelLower === 'diagnostic') {
      currentCategory = currentCategory.split(' ')[0] + ' diagnostic';
      continue;
    } else if (labelLower === 'rehabilitation' || labelLower === 'rehab') {
      currentCategory = currentCategory.split(' ')[0] + ' rehab';
      continue;
    }
    
    // Skip total/summary rows
    if (labelLower.indexOf('total') !== -1) continue;
    if (labelLower.indexOf('clinic name') !== -1) continue;
    
    // This is an activity row — extract hours for each date
    for (var di in dateColumns) {
      var dc = dateColumns[di];
      var observed = parseFloat(data[r][dc.obsCol]) || 0;
      var tested = parseFloat(data[r][dc.testCol]) || 0;
      
      if (observed > 0 || tested > 0) {
        result.entries.push({
          date: dc.date,
          dateIndex: di,
          category: currentCategory,
          activity: label,
          observed: observed,
          tested: tested
        });
      }
    }
  }
  
  return result;
}

/**
 * Maps an activity to a clinic type based on its category.
 */
function categoriseActivity(category, activity) {
  var cat = category.toLowerCase();
  if (cat.indexOf('adult') !== -1 && cat.indexOf('diagnostic') !== -1) return 'Adult diagnostic';
  if (cat.indexOf('adult') !== -1 && cat.indexOf('rehab') !== -1) return 'Adult rehab';
  if (cat.indexOf('paediatric') !== -1 && cat.indexOf('diagnostic') !== -1) return 'Paediatric diagnostic';
  if (cat.indexOf('paediatric') !== -1 && cat.indexOf('rehab') !== -1) return 'Paediatric rehab';
  if (cat.indexOf('other') !== -1) return 'Other';
  
  // Fallback: try to guess from activity name
  var act = activity.toLowerCase();
  if (act.indexOf('vra') !== -1 || act.indexOf('play') !== -1 || act.indexOf('ktt') !== -1 ||
      act.indexOf('newborn') !== -1 || act.indexOf('boa') !== -1) return 'Paediatric diagnostic';
  if (act.indexOf('hearing aid') !== -1 || act.indexOf('rem') !== -1 ||
      act.indexOf('fitting') !== -1 || act.indexOf('impression') !== -1) return 'Adult rehab';
  if (act.indexOf('observation') !== -1 || act.indexOf('simulation') !== -1) return 'Other';
  
  return 'Adult diagnostic'; // Default
}

/**
 * Checks if a value looks like a date.
 */
function isDateLike(val) {
  if (!val) return false;
  if (typeof val === 'number' && val > 40000 && val < 50000) return true; // Excel serial date
  if (typeof val === 'string' && !isNaN(Date.parse(val))) return true;
  return false;
}

/**
 * Formats a date for use in session IDs.
 */
function formatDateForId(date) {
  if (!(date instanceof Date)) date = new Date(date);
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, '0');
  var d = String(date.getDate()).padStart(2, '0');
  return y + m + d;
}

/**
 * Logs each import for audit purposes.
 */
function logImport(studentId, clinicName, term, rowCount, observed, tested, fileId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('ImportLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('ImportLog');
    logSheet.appendRow(['Timestamp', 'StudentID', 'PlacementClinic', 'Term',
                        'RowsImported', 'TotalObserved', 'TotalTested', 'FileID', 'ImportedBy']);
  }
  logSheet.appendRow([
    new Date(), studentId, clinicName, term,
    rowCount, observed, tested, fileId,
    Session.getActiveUser().getEmail()
  ]);
}

/**
 * Shows the import log.
 */
function showImportLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('ImportLog');
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  } else {
    SpreadsheetApp.getUi().alert('No imports have been logged yet.');
  }
}

/**
 * Lists sessions from placements that don't have reflections confirmed.
 */
function listPendingReviews() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoursTab = ss.getSheetByName('Hours');
  if (!hoursTab) {
    SpreadsheetApp.getUi().alert('No Hours tab found.');
    return;
  }
  
  var data = hoursTab.getDataRange().getValues();
  var pending = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] === 'Placement' && data[i][13] !== true) {
      var key = data[i][1] + '|' + data[i][5]; // StudentID|Date
      if (pending.indexOf(key) === -1) pending.push(key);
    }
  }
  
  if (pending.length === 0) {
    SpreadsheetApp.getUi().alert('All placement sessions have confirmed reflections.');
  } else {
    SpreadsheetApp.getUi().alert(
      'Placement sessions pending reflection confirmation: ' + pending.length + '\n\n' +
      pending.slice(0, 20).map(function(p) {
        var parts = p.split('|');
        return parts[0] + ' — ' + parts[1];
      }).join('\n') +
      (pending.length > 20 ? '\n... and ' + (pending.length - 20) + ' more' : '')
    );
  }
}

/**
 * Marks placement hours as reflection-confirmed for a given session.
 * Call this when a student submits their PebblePad reflection.
 * 
 * @param {string} sessionId - The session ID (e.g. PLC-2025001-20250615)
 */
function confirmPlacementReflection(sessionId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoursTab = ss.getSheetByName('Hours');
  if (!hoursTab) return;
  
  var data = hoursTab.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][12] === sessionId) {
      hoursTab.getRange(i + 1, 14).setValue(true); // Column N = ReflectionConfirmed
    }
  }
}
