/******************************************
* Call Schedule Manager
* @version: 1.1
* @author: Naman Jindal (nj.itprof@gmail.com)
******************************************/

var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';
var TAB_NAME = 'Dashboard Urls';


function main() {
  var input = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  input.shift();
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var hour = NOW.getHours();
  
  var shouldRun = false;
  
  for(var k in input) {
    var SETTINGS = {};
    SETTINGS.LABEL = input[k][0];
    SETTINGS.REPORT_URL = input[k][1];
    if(SETTINGS.LABEL == 'Monique') { 
      shouldRun = true;
      continue;
    }
    if(!shouldRun) { continue; }
    
    runScript(SETTINGS);
    
  }
}

function runScript(SETTINGS) {
  var HEADER = ['Call Extension Id','Account Name','Phone Number',
                'Monday','Tuesday','Wednesday','Thursday','Friday',
                'Saturday','Sunday'];
  
  var sheetName = 'Call Schedule Manager';
  
 var URL = '';
  var ss = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL);
  var tab = ss.getSheetByName('Report Links');
  var data = tab.getDataRange().getValues();
  data.shift();
  for(var z in data) {
    if(data[z][2] == 'CALL_SCHEDULE_MANAGER_URL') {
      if(data[z][1]) { 
        URL = data[z][1];
      }
    }
  }
  
  if(!URL) { return; }
  
  var ss = SpreadsheetApp.openByUrl(URL);
  var sheet = ss.getSheets()[0];
  sheet.hideColumns(1);
  sheet.getDataRange().setFontFamily('Calibri');
  
  Logger.log('Running for: ' + SETTINGS.LABEL);
  
  var actionFlag = sheet.getRange('C1').getValue();
  if(!actionFlag || actionFlag == 'Pause') { return; }
  
  var existingData = sheet.getDataRange().getValues();
  existingData.shift();
  var rowHeader = existingData.shift();
  
  var scheduleMap = {};
  if(actionFlag == 'Apply Changes') {
    for(var k in existingData) {
      if(!scheduleMap[existingData[k][1]]) {
        scheduleMap[existingData[k][1]] = {};
      }
      
      scheduleMap[existingData[k][1]][existingData[k][0]] = [];
      
      for(var j in existingData[k]) {
        if(j < 3 || !existingData[k][j]) { continue; }
        var parts = existingData[k][j].split('-');
        
        var startHour = parseInt(parts[0].trim().split(':')[0],10);
        var startMin = parseInt(parts[0].trim().split(':')[1],10);
        var endHour = parseInt(parts[1].trim().split(':')[0],10)
        var endMin = parseInt(parts[1].trim().split(':')[1],10);
        
        scheduleMap[existingData[k][1]][existingData[k][0]].push({ 
          'dayOfWeek': rowHeader[j].toUpperCase(),
          'startHour': startHour,
          'startMinute': startMin,
          'endHour': endHour,
          'endMinute': endMin
        }); 
      }
    }
  }
  
  var data = [];
  
  var accountIter = MccApp.accounts()
  .withCondition('LabelNames CONTAINS "'+SETTINGS.LABEL+'"')
 // .withCondition('Name = "Blue Sky Resumes"')
  .orderBy('Name ASC')
  .get();
  
  while(accountIter.hasNext()) { 
    var account = accountIter.next();
    MccApp.select(account);
    if(!AdWordsApp.currentAccount().getName()) { return ''; }
    
    if(actionFlag == 'Apply Changes') {
      applyChanges(scheduleMap, data);
    } else {
      exportData(data);
    }
  }
  
  sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  if(data.length > 0) {
    sheet.getRange(3,1,data.length,data[0].length).setValues(data);
  }
  
  sheet.getRange(1, 3, sheet.getLastRow(), sheet.getLastColumn()-3).setHorizontalAlignment('center');
  sheet.setFrozenRows(2);
  sheet.getRange('C1').setValue('Export');
  sheet.getDataRange().setFontFamily('Calibri');
  
  var col = 1;
  while(col <= sheet.getLastColumn()) {
   sheet.autoResizeColumn(col);
    col++;
  }
  
  deleteExtraCols(sheet);
}

function deleteExtraCols(sheet) {
  if((sheet.getMaxColumns() - sheet.getLastColumn()) > 0) {
    sheet.deleteColumns(sheet.getLastColumn()+1, sheet.getMaxColumns() - sheet.getLastColumn());
  }
}


function applyChanges(scheduleMap, data) {
  
  var accName = AdWordsApp.currentAccount().getName();
  var schedules = scheduleMap[accName];

  var indexMap = { 'MONDAY': 3, 'TUESDAY': 4, 'WEDNESDAY': 5, 'THURSDAY': 6, 'FRIDAY': 7, 'SATURDAY': 8, 'SUNDAY': 9 }
  var callIter = AdWordsApp.extensions().phoneNumbers().get();
  while(callIter.hasNext()) {
    var call = callIter.next();
    if(schedules) { 
      var newSchedules = schedules[call.getId()];
      if(newSchedules) {
        call.setSchedules(newSchedules);
      }
    }
    
    var phoneNumber = call.isMobilePreferred() ? call.getPhoneNumber() + ' (mobile)' : call.getPhoneNumber();
    var row = [call.getId(), accName, phoneNumber, '', '', '', '', '', '', ''];
    
    var existingSchedules = call.getSchedules();
    for(var k in existingSchedules) {
      var index = indexMap[existingSchedules[k].getDayOfWeek()];
      
      var startTime = existingSchedules[k].getStartHour() + ':' + ('0'+existingSchedules[k].getStartMinute()).slice(-2);
      var endTime = existingSchedules[k].getEndHour() + ':' + ('0'+existingSchedules[k].getEndMinute()).slice(-2);
      row[index] = startTime + ' - ' + endTime;
    }
    
    data.push(row);
  }
}

function exportData(data) {
  var indexMap = { 'MONDAY': 3, 'TUESDAY': 4, 'WEDNESDAY': 5, 'THURSDAY': 6, 'FRIDAY': 7, 'SATURDAY': 8, 'SUNDAY': 9 }
  var accName = AdWordsApp.currentAccount().getName();
  var callIter = AdWordsApp.extensions().phoneNumbers().get();
  while(callIter.hasNext()) {
    var call = callIter.next();
    var phoneNumber = call.isMobilePreferred() ? call.getPhoneNumber() + ' (mobile)' : call.getPhoneNumber();
    var row = [call.getId(), accName, phoneNumber, '', '', '', '', '', '', ''];
    
    var schedules = call.getSchedules();
    for(var k in schedules) {
      var index = indexMap[schedules[k].getDayOfWeek()];
      
      var startTime = schedules[k].getStartHour() + ':' + ('0'+schedules[k].getStartMinute()).slice(-2);
      var endTime = schedules[k].getEndHour() + ':' + ('0'+schedules[k].getEndMinute()).slice(-2);
      row[index] = startTime + ' - ' + endTime;
    }
    
    data.push(row);
  }
}