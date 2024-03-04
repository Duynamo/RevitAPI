function createTriggers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("checkStatus").forSpreadsheet(sheet).onChange().create();
}

function checkStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeD = sheet.getRange('D10:D1000');
  var valuesD = rangeD.getValues();
  var rangeJ = sheet.getRange('J10:J1000');
  var valuesJ = rangeJ.getValues();
  var rangeK = sheet.getRange('K10:K1000');
  var valuesK = rangeK.getValues();
  var rangeL = sheet.getRange('L10:L1000');
  var valuesL = rangeL.getValues();

  var range = sheet.getRange('D10:D1000');
  var output = [];
  var colors = [];

  for (var i = 0; i < valuesJ.length; i++) {
    var valueD = valuesD[i][0];
    var valueJ = valuesJ[i][0];
    var valueK = valuesK[i][0];
    var valueL = valuesL[i][0]; // Lấy giá trị tại vị trí i trong mảng valuesL
    var color = '';

    if (valueJ == 0 && valueK == 0) {
      output.push(['']);
    } else if (valueJ == valueK && valueJ < 1) {
      output.push(['進行中']);
      color = '#ffA500';
    } else if (valueJ > valueK) {
      output.push(['進行中']);
      color = '#ff0000';
    } else if (valueJ == valueK && valueJ == 1) {
      output.push(['完了']);
      color = '#00ff00';
    } else if (valueK > valueJ && valueK < 1) {
      output.push(['進行中']);
      color = '#ffA500';
    } else if (valueJ < valueK && valueK == 1) {
      output.push(['完了']);
      color = '#00ff00';
    } else if (valueL.includes('休止')) { // Kiểm tra giá trị của valuesL
      output.push(['休止']); // Ghi giá trị '休止'
      color = '#d9d9d9'; // Tô màu nền xám
    } else {
      output.push(['']);
    }
    colors.push([color]);
  }

  range.setValues(output);
  range.setBackgrounds(colors);
}
