function CopyTimeData() {
  //Define source and summary sheets
  let spreadSheet =  SpreadsheetApp.getActiveSpreadsheet();
  let sourceSheet = spreadSheet.getSheetByName('TimeSheet');
  let summarySheet = spreadSheet.getSheetByName(sourceSheet.getRange('C2').getValue());

  //Popup message 
  let ui = SpreadsheetApp.getUi();
  let button = ui.alert("Nhấn OK để lưu dữ liệu, nhấn Cancel để chỉnh sửa lại dữ liệu.", ui.ButtonSet.OK_CANCEL);

  if(button == "OK")
  {
    //Count exist rows in summary sheet
    var summaryData = summarySheet.getRange('A:A').getValues().flat();
    var count = summaryData.filter(Boolean).length;

    //Copy data from source to summary
    let sourceRange = sourceSheet.getRange("A6:G33");
    let sourceContentRange = sourceSheet.getRange("E6:G33");
    let sourceValue  = sourceRange.getValues();

    var currentDate = sourceSheet.getRange("C3").getValue();
    var result = !summaryData.toString().includes(currentDate.toString()); // Check if data existed

    if(result)
    {
      summarySheet.getRange(count+1,1,28,1).setValue(currentDate);
      let targetRange = summarySheet.getRange(count+1,2,28,7);
      targetRange.setValues(sourceValue);
      sourceContentRange.clearContent()

      //Sort summary sheet by date
      summarySheet.getRange('A:H').sort(1)
      ui.alert("Dữ liệu đã được lưu.", ui.ButtonSet.OK)
    }
    else(ui.alert("Dữ liệu đã tồn tại.Vui lòng chọn chỉnh sửa.", ui.ButtonSet.OK))

  }


}

function EditTimeData() {
  //Define source and summary sheets
  let spreadSheet =  SpreadsheetApp.getActiveSpreadsheet();
  let sourceSheet = spreadSheet.getSheetByName('TimeSheet');
  let summarySheet = spreadSheet.getSheetByName(sourceSheet.getRange('C2').getValue());

  //Popup message 
  let ui = SpreadsheetApp.getUi();
  let button = ui.alert("Nhấn OK để lưu dữ liệu, nhấn Cancel để chỉnh sửa lại dữ liệu.", ui.ButtonSet.OK_CANCEL);

  if(button == "OK")
  {
    //Count exist rows in summary sheet
    var summaryData = summarySheet.getRange('A:A').getValues().flat();
    var count = summaryData.filter(Boolean).length;

    //Edit data
    let sourceRange = sourceSheet.getRange("A6:G33");
    let sourceContentRange = sourceSheet.getRange("E6:G33");
    let sourceValue  = sourceRange.getValues();

    var editDate = sourceSheet.getRange("L5").getValue();
    var result = !summaryData.toString().includes(editDate.toString()); // Check if data existed

    if(result)
    {
      summarySheet.getRange(count+1,1,28,1).setValue(editDate);
      let targetRange = summarySheet.getRange(count+1,2,28,7);
      targetRange.setValues(sourceValue);
    }
    else{
      for(var i =1; i<=count; i++)
      {
        if(editDate.toString()==summaryData[i].toString()){
          e = i;
          break
        }
      }
      summarySheet.getRange(e,1,28,1).setValue(editDate);
      let targetRange = summarySheet.getRange(e,2,28,7);
      targetRange.setValues(sourceValue);
    }

    sourceContentRange.clearContent()

    //Sort summary sheet by date
    summarySheet.getRange('A:H').sort(1)
    ui.alert("Dữ liệu đã được lưu.", ui.ButtonSet.OK)

  }
}