function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

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
