function doPost(e) {
  try {
    var labname = e.parameter.lab;
    if (!labname) throw new Error("Missing 'lab' parameter");

    var student = e.parameter.student;
    if (!student) throw new Error("Missing 'student' parameter");

    let app = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1WYY9cxLVQUTfrtWBSy_nTw8g2pvodk5ET--tlBcWkbs/edit?gid=0#gid=0");
    let sheet = app.getSheetByName(labname);
    if (!sheet) throw new Error("Sheet not found");

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] == student && data[i][2] == labname) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: 'คุณได้ส่งข้อมูลแล้วใน Lab นี้'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ตรวจสอบว่าได้ส่งไฟล์หรือไม่
    if (!e.postData.contents) throw new Error("No data received");

    var files = e.parameter;
    var links = [];

    for (var i = 0; i < Object.keys(files).length; i++) {
      var fileData = files["file" + i];
      var decodedFile = Utilities.base64Decode(fileData);
      var blob = Utilities.newBlob(decodedFile);
      var newFile = DriveApp.createFile(blob);
      var link = newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW).getDownloadUrl();
      links.push(link);
    }

    // เพิ่มลิงก์ไฟล์ไปใน Google Sheets
    let lr = sheet.getLastRow();
    links.forEach(function(link, index) {
      sheet.getRange(lr, index + 4).setFormula(`=IMAGE("${link}")`); // แทรกสูตร IMAGE ในคอลัมน์ถัดไป
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Images uploaded successfully",
      imageUrls: links
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
