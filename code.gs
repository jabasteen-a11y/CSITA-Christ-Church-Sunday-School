function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
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
    data.sunday_school_class,
    data.academic_year,
    data.teacher_name,
    data.teacher_contact,
    data.remarks
  ]);

  return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
