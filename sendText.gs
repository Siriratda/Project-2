function doGet(e) {
  var callback = e.parameter.callback;
  var name = e.parameter.name;
  var student = e.parameter.student;
  var labname = e.parameter.lab;
  var score = e.parameter.score;

  let app = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1WYY9cxLVQUTfrtWBSy_nTw8g2pvodk5ET--tlBcWkbs/edit?gid=0#gid=0");
  var sheet = app.getSheetByName(labname);
  var data = sheet.getDataRange().getValues();

  // เช็คว่ามีข้อมูลซ้ำใน Google Sheets
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == student && data[i][2] == labname) {
      var response = {
        result: 'error',
        message: 'คุณได้ส่งข้อมูลแล้วใน Lab นี้'
      };
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(response) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }

  // หากไม่มีข้อมูลซ้ำ บันทึกข้อมูลลงใน Google Sheets
  sheet.appendRow([name, student, labname, score]);

  var response = {
    result: 'success',
    name: name,
    student: student,
    lab: labname,
    score: score
  };

  return ContentService.createTextOutput(callback + '(' + JSON.stringify(response) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
