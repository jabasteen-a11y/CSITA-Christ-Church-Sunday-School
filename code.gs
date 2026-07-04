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

  // Count existing students in same class
  var allData = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][2] === data.sunday_school_class) {
      count++;
    }
  }
  var serial = count + 1;
  var regId = serial + gender + '-' + classCode + '-' + year;

  // ── Append row ──────────────────────────────────────────────
  sheet.appendRow([
    regId,
    new Date(),
    data.sunday_school_class,
    data.student_name,
    data.dob,
    data.gender,
    data.school_standard,
    data.school_name,
    data.address,
    data.father_name,
    data.father_profession,
    data.father_contact,
    data.mother_name,
    data.mother_profession,
    data.mother_contact,
    data.academic_year,
    data.teacher_name,
    data.remarks
  ]);

  // ── Send Registration ID back to website ────────────────────
  return ContentService.createTextOutput(JSON.stringify({
    result: 'success',
    reg_id: regId,
    student_name: data.student_name,
    sunday_school_class: data.sunday_school_class
  })).setMimeType(ContentService.MimeType.JSON);
}
