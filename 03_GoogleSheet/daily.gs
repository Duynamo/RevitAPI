function onOpen(){
  let spreadSheet =  SpreadsheetApp.getActiveSpreadsheet();
  let sourceSheet = spreadSheet.getSheetByName('FVC007');
  let dailySheet = spreadSheet.getSheetByName('Daily');

  let dataSource = sourceSheet.getDataRange().getValues();
  let days = dataSource.map(x => x[0]).map(item => item.toString());
  let daySource = [...new Set(days)];
  let projectField = spreadSheet.getSheetByName('Field').getRange("B2:C100").getValues().filter(item => item[0] != '');
  //console.log(dataSource.slice(0,28));
  for(j = 0; j < daySource.length; j++){
    let dataRange = dataSource.slice(j*28, (j+1)*28);
    let day = dataRange[0][0];
    let data = [];
    let dataLine = row => {if(row[0].toString() === day.toString()){data.push(row)}};
    //console.log(day);
    dataRange.forEach(dataLine);
    workDetails = filterWork(data, projectField);
    let dailyWorks = [];
    workDetails.forEach(item => {dailyWorks.push(joinWorkDetail(item))});
    dailySheet.getRange(j+1, 1).setValue(dataRange[0][0]);
    dailySheet.getRange(j+1, 2).setValue(dailyWorks.join('\n'));
  }
  dailySheet.getRange('A:B').sort(1);
}

function filterWork(data, projectField){
  let work = removeIfNextEqual(data.map(x => x[6])).filter(item => (item != "" && item!= "MT00_朝礼・終礼"));
  startTimes = [];
  endTimes = [];
  projects = []
  workDetails = []

  for(i = 0; i < work.length; i++){
    task = work[i]
    let startTime = data.map(x =>{if(x[6] === task) {return[x[1],x[2]]}}).filter(item => item!= null)[0];
    let endTime = data.map(x =>{if(x[6] === task) {return[x[3],x[4]]}}).filter(item => item!= null).slice(-1)[0];
    let projcode = data.map(x =>{if(x[6] === task) {return[x[5]]}}).filter(item => item!= null)[0];
    let project = projectField.filter(x => x[0] === projcode.toString())[0][1];
    let workDetail = [data.map(x =>{if(x[6] === task) {return[x[7]]}}).filter(item => item!= null).
      filter(item => item != "")].join();
    startTimes.push(startTime);
    endTimes.push(endTime);
    projects.push(project);
    workDetails.push(workDetail);
  }
  return work.map(function(task, index) {
    return [startTimes[index], endTimes[index] , projects[index], task.toString().slice(5),  workDetails[index]];
    })
}

function joinWorkDetail(workDetail){
  return "【" + workDetail[0][0] + ":" + workDetail[0][1] + "-" + workDetail[1][0] + ":" + workDetail[1][1] 
          + ", " + workDetail[2] + ", " + workDetail[3] + "】" + "\n"
          + "・" + workDetail[4]
}

function removeIfNextEqual(list) {
    list = list.filter(item => item!="")
    let i = 0;
    while (i < list.length - 1) {
        if (list[i] === list[i + 1]) {
            list.splice(i, 1);
        } else {
            i++;
        }
    }
    return list;
}

