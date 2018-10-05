
var MASTER_DASHBOARD = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';
var DASHBOARD_URLS_TAB_NAME = 'Dashboard Urls';
var MANAGEMENT_URLS_TAB_NAME = 'Management Urls';
var TAB_MAP = { 
  'Label Ads By Performance': 1, 'Daily Snapshot': 1, 
  'Bid To Position': 1, 'Zero Clicks Alert': 1,
  'Mobile Bids and Placements': 1
} 

var HOURS_TO_RUN = [0,4,8,12,16,18,22];

function main() {
  try {
    runForType(DASHBOARD_URLS_TAB_NAME);
  } catch(E) {
    
  }
  
  try {
    runForType('Franchisee Dashboard Urls');
  } catch(E) {
    
  }
  
  try {
    runForType('External Dashboard Urls')
  } catch(E) {
    
  }
}

function runForType(TAB_NAME) {
  //return;
  var names = {};
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'),10);
  //if(HOURS_TO_RUN.indexOf(hour) < 0) { return; }
  
  var commonInputs = {};
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD).getSheetByName(TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var urlMap = {};
  for(var k in data) {
    urlMap[data[k][0]] = data[k][2];
  }
  
  for(var k in data) {
    cleanDashboard(names, data[k][0], data[k][1], commonInputs, urlMap);
    
    if(data[k][0] == 'New Business') {
      updateCommonInputs(commonInputs, data[k][1]); 
    }
  }
  
  cleanMasterManagementSheets(names);
}

function updateCommonInputs(map, URL) {
  var ss = SpreadsheetApp.openByUrl(URL);
  var tab = ss.getSheetByName('Account Inputs');
  
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  
  for(var z in data) {
    if(!map[data[z][0]]) { continue; }
    tab.getRange(parseInt(z,10)+3, 2, 1, map[data[z][0]].length).setValues([map[data[z][0]]]);
  }
}

function cleanDashboard(names,LABEL,URL,commonInputs,urlMap) {

  
  var ss = SpreadsheetApp.openByUrl(URL);
  var managementReportSS = ss.getSheetByName('Report Links').getRange(4,2).getValue();
  
  var tab = ss.getSheetByName('Management Report');
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var ssUrl = managementReportSS.split('edit')[0];
  if(!names[ssUrl]) {
    names[ssUrl] = {};
  }
  
  for(var x in data) {
    if(!data[x][0]) { continue; }
    
    names[ssUrl][data[x][0]] = 1;
    if(data[x][1] && urlMap[data[x][1]]) {
      var tempUrl = urlMap[data[x][1]].split('edit')[0];
      if(!names[tempUrl]) {
        names[tempUrl] = {};
      }    
      names[tempUrl][data[x][0]] = 1;
    }
    
    if(data[x][2]) {
      var tempUrl = data[x][2].split('edit')[0];
      if(!names[tempUrl]) {
        names[tempUrl] = {};
      }    
      names[tempUrl][data[x][0]] = 1;
    }
  }
  //Logger.log(LABEL);
  //Logger.log(Object.keys(names))
  if(LABEL == 'New Business') { return; }
  
  var data = ss.getSheetByName('Account Inputs').getDataRange().getValues();
  data.shift();
  data.shift();
  
  for(var z in data) {
    if(commonInputs[data[z][0]] && !data[z][2]) { continue; }
    commonInputs[data[z][0]] = [data[z][1], data[z][2], data[z][3], data[z][4], data[z][5], data[z][6], data[z][7]]; 
  }
  
  
}

function cleanMasterManagementSheets(names) {
  
  for(var url in names) {
    var mUrl = url + 'edit';
    Logger.log(mUrl);
    var ss = SpreadsheetApp.openByUrl(mUrl);
    var sheet = ss.getSheetByName('Daily Report');
    // var managementSpreadsheet  = SpreadsheetApp.openByUrl(managementReportSS);
    var data = sheet.getDataRange().getValues();
    
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    
    
    
    for(var k = data.length-1; k>=0; k--) {
      if(!data[k][1]) { continue; }
      
      if(!names[url][data[k][1]]) {
        var row = parseInt(k,10)+6;
        //Logger.log(mUrl + '::' + row);
        sheet.deleteRow(row)
      }
    }
    
    var sheet = ss.getSheetByName('Weekly & Monthly Report');
    var data = sheet.getDataRange().getValues();
    
    data.shift();
    data.shift();
    
    for(var k = data.length-1; k>=0; k--) {
      if(!names[url][data[k][0]]) {
        var row = parseInt(k,10)+3;
        //Logger.log(row);
        sheet.deleteRow(row)
      }
    }
    
    var sheet = ss.getSheetByName('Spends Report');
    if(sheet) {
      var data = sheet.getDataRange().getValues();
      
      data.shift();
      data.shift();
      
      for(var k = data.length-1; k>=0; k--) {
        if(!data[k][1]) { continue; }
        
        if(!names[url][data[k][1]]) {
          var row = parseInt(k,10)+3;
          //Logger.log(row);
          sheet.deleteRow(row)
        }
      }
    }
    
    var sheet = ss.getSheetByName('Daily Budget Report');
    if(sheet) {
      var data = sheet.getDataRange().getValues();
      
      data.shift();
      data.shift();
      
      for(var k = data.length-1; k>=0; k--) {
        if(!data[k][1]) { continue; }
        
        if(!names[url][data[k][1]]) {
          var row = parseInt(k,10)+3;
          //Logger.log(row);
          sheet.deleteRow(row)
        }
      }
    }
    
    var sheet =  ss.getSheetByName('KPI Report');
    if(sheet) {
      var inputData = sheet.getDataRange().getValues();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      
      var toDelete = [];
      for(var l in inputData) {
        if(!inputData[l][1]) { continue; }
        if(!inputData[l][1]) { continue; }
        if(!names[url][inputData[l][1]]) {
          toDelete.push(parseInt(l,10)+6); 
        }
      }
      
      try {
        for(var l=toDelete.length-1; l >= 0; l--) {
          sheet.deleteRow(toDelete[l]);
        }
      } catch(ex) {
        Logger.log(ex); 
      }
    }
    
    
    var sheet =  ss.getSheetByName('Ecommerce Report');
    if(sheet) {
      var inputData = sheet.getDataRange().getValues();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      
      var toDelete = [];
      for(var l in inputData) {
        if(!inputData[l][0]) { continue; }
        if(!names[url][inputData[l][0]]) {
          toDelete.push(parseInt(l,10)+5); 
        }
      }
      
      for(var l=toDelete.length-1; l >= 0; l--) {
        sheet.deleteRow(toDelete[l]);
      }
    }
    
    var sheet =  ss.getSheetByName('Conversion Report');
    if(sheet) {
      var inputData = sheet.getDataRange().getValues();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      inputData.shift();
      
      var toDelete = [];
      for(var l in inputData) {
        if(!inputData[l][0]) { continue; }
        if(!names[url][inputData[l][0]]) {
          toDelete.push(parseInt(l,10)+5); 
        }
      }
      
      try {
        for(var l=toDelete.length-1; l >= 0; l--) {
          sheet.deleteRow(toDelete[l]);
        }
      } catch(ex) {
        Logger.log(ex); 
      }
    } 
    
    var sheet =  ss.getSheetByName('Email Templates');
    if(sheet) {
      var inputData = sheet.getDataRange().getValues();
      inputData.shift();
      
      var toDelete = [];
      for(var l in inputData) {
        if(!inputData[l][0]) { continue; }
        if(!names[url][inputData[l][0]]) {
          toDelete.push(parseInt(l,10)+2); 
        }
      }
      
      try {
        for(var l=toDelete.length-1; l >= 0; l--) {
          sheet.deleteRow(toDelete[l]);
        }
      } catch(ex) {
        Logger.log(ex); 
      }
    }
  }
}