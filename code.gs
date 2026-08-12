// ── Column reference (1-indexed for getRange, 0-indexed for arrays) ──
// A regId | B timestamp | C name | D dob | E gender | F school_standard
// G school_name | H address | I father_name | J father_profession
// K father_contact | L mother_name | M mother_profession | N mother_contact
// O sunday_school_class | P academic_year | Q teacher_name | R remarks

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (data.action === 'updateStudent') {
    return handleUpdateStudent(sheet, data);
  }
  if (data.action === 'deleteStudent') {
    return handleDeleteStudent(sheet, data);
  }
  return handleNewRegistration(sheet, data);
}

function handleNewRegistration(sheet, data) {
  // ── Class code map ──────────────────────────────────────────
  var CLASS_CODES = {
    'Tiny Tots (Pre-Beginners)': 'TT',
    'Beginners':                 'BG',
    'Primary Boys':              'PB',
    'Primary Girls':             'PG',
    'Junior Boys':               'JB',
    'Junior Girls':              'JG',
    'Inter-Boys':                'IB',
    'Inter-Girls':               'IG',
    'Senior Boys':               'SB',
    'Senior Girls-A':            'SGA',
    'Senior Girls-B':            'SGB'
  };

  // ── Generate Registration ID ────────────────────────────────
  var classCode = CLASS_CODES[data.sunday_school_class] || 'XX';
  var gender    = (data.gender === 'Male') ? 'M' : 'F';
  var year      = data.academic_year || '2026-27';

  // Count existing students in same class (Class is column O = index 14)
  var allData = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][14] === data.sunday_school_class) {
      count++;
    }
  }
  var serial = count + 1;
  var regId = serial + gender + '-' + classCode + '-' + year;

  // ── Append row ──────────────────────────────────────────────
  sheet.appendRow([
    regId,                       // A: Registration ID
    new Date(),                  // B: Timestamp
    data.student_name,           // C: Student Name
    data.dob,                    // D: DOB
    data.gender,                 // E: Gender
    data.school_standard,        // F: School Standard
    data.school_name,            // G: School Name
    data.address,                // H: Address
    data.father_name,            // I: Father Name
    data.father_profession,      // J: Father Profession
    data.father_contact,         // K: Father Contact
    data.mother_name,            // L: Mother Name
    data.mother_profession,      // M: Mother Profession
    data.mother_contact,         // N: Mother Contact
    data.sunday_school_class,    // O: Class
    data.academic_year,          // P: Academic Year
    data.teacher_name,           // Q: Teacher Name
    data.remarks                 // R: Remarks
  ]);

  // ── Send Registration ID back ───────────────────────────────
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    reg_id: regId,
    student_name: data.student_name,
    sunday_school_class: data.sunday_school_class
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── NEW: update an existing student's details, found by Registration ID ──
// Note: editing "Class" here does NOT regenerate the Registration ID —
// the ID stays fixed once created, even if the class value changes.
function handleUpdateStudent(sheet, data) {
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.regId) {
      var rowNum = i + 1; // sheet rows are 1-indexed

      sheet.getRange(rowNum, 3).setValue(data.student_name);       // C
      sheet.getRange(rowNum, 4).setValue(data.dob);                // D
      sheet.getRange(rowNum, 5).setValue(data.gender);             // E
      sheet.getRange(rowNum, 6).setValue(data.school_standard);    // F
      sheet.getRange(rowNum, 7).setValue(data.school_name);        // G
      sheet.getRange(rowNum, 8).setValue(data.address);            // H
      sheet.getRange(rowNum, 9).setValue(data.father_name);        // I
      sheet.getRange(rowNum, 10).setValue(data.father_profession); // J
      sheet.getRange(rowNum, 11).setValue(data.father_contact);    // K
      sheet.getRange(rowNum, 12).setValue(data.mother_name);       // L
      sheet.getRange(rowNum, 13).setValue(data.mother_profession); // M
      sheet.getRange(rowNum, 14).setValue(data.mother_contact);    // N
      sheet.getRange(rowNum, 15).setValue(data.sunday_school_class); // O
      sheet.getRange(rowNum, 16).setValue(data.academic_year);     // P
      sheet.getRange(rowNum, 17).setValue(data.teacher_name);      // Q
      sheet.getRange(rowNum, 18).setValue(data.remarks);           // R

      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        message: 'Student updated'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    result: 'error',
    message: 'Student not found for regId: ' + data.regId
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── NEW: delete a student row, found by Registration ID ──────────────────
// Uses the same sheet reference and column-A lookup convention as
// handleUpdateStudent/handleNewRegistration above (regId is always column A).
function handleDeleteStudent(sheet, data) {
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.regId) {
      sheet.deleteRow(i + 1); // +1 because sheet rows are 1-indexed

      return ContentService.createTextOutput(JSON.stringify({
        result: 'success',
        deletedRegId: data.regId
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    result: 'error',
    message: 'Registration ID not found: ' + data.regId
  })).setMimeType(ContentService.MimeType.JSON);
}

// ── doGet handles read-only queries: the admin students list ──
function doGet(e) {
  var action = e.parameter.action;

  if (action === 'listStudents') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allData = sheet.getDataRange().getValues();
    var students = [];

    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      if (!row[0]) continue; // skip blank rows

      students.push({
        regId: row[0],                  // A
        student_name: row[2],           // C
        dob: formatDobForField(row[3]), // D
        gender: row[4],                 // E
        school_standard: row[5],        // F
        school_name: row[6],            // G
        address: row[7],                // H
        father_name: row[8],            // I
        father_profession: row[9],      // J
        father_contact: row[10],        // K
        mother_name: row[11],           // L
        mother_profession: row[12],     // M
        mother_contact: row[13],        // N
        sunday_school_class: row[14],   // O
        academic_year: row[15],         // P
        teacher_name: row[16],          // Q
        remarks: row[17]                // R
      });
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      students: students
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── Used by the Harvest Festival competition registration page ──
  if (action === 'listByClass') {
    var className = e.parameter.class;
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allData2 = sheet2.getDataRange().getValues();
    var matches = [];

    for (var j = 1; j < allData2.length; j++) {
      var row2 = allData2[j];
      if (!row2[0]) continue;
      if (row2[14] === className) { // O: Class
        matches.push({
          regId: row2[0],           // A
          student_name: row2[2],    // C
          school_standard: row2[5]  // F
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      students: matches
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    result: 'error',
    message: 'Unknown or missing action'
  })).setMimeType(ContentService.MimeType.JSON);
}

// Sheet may store DOB as a JS Date object; the edit form's <input type="date">
// needs it as yyyy-MM-dd text, otherwise it comes through as an object/string mismatch.
function formatDobForField(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}
