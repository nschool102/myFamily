function doPost(e) {
  var requestData = JSON.parse(e.postData.contents);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TRANSACTIONS");
  
  if (requestData.action === "syncTransactions") {
    requestData.data.forEach(function(item) {
      // Định dạng ngày giờ chuẩn dễ đọc cho Google Sheets
      var formattedDate = Utilities.formatDate(new Date(item.timestamp), "GMT+7", "dd/MM/yyyy HH:mm");
      
      // Ghi chuẩn xác theo cấu trúc 5 cột: A (TIMESTAMP) | B (TYPE) | C (SUBTYPE) | D (SỐ TIỀN) | E (GHI CHÚ)
      sheet.appendRow([
        formattedDate,
        item.type,
        item.subtype,
        item.amount,
        item.note
      ]);
    });
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e.parameter.action === "getFamily") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FAMILY");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var member = {};
      headers.forEach(function(header, index) {
        // Map chuẩn xác tên cột từ dữ liệu sheet FAMILY
        member[header.toLowerCase().replace(/[\s:]/g, "")] = row[index];
      });
      result.push(member);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}