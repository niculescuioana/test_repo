
var MASTER_DASHBOARD = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';
var DASHBOARD_URLS_TAB_NAME = 'Dashboard Urls';

var TAB_MAP = { 
  'Account Inputs': 1, 'Label Ads By Performance': 1, 'Daily Snapshot': 1, 
  'Bid To Position': 1, 'Zero Clicks Alert': 1
} 

function main() {
  
  try {
    cleanPushDashboards();
  } catch(ee) {
    
  }
  
  try {
    cleanFranchiseeDashboards();
  } catch(ee) {
    
  }
  
  try {
    cleanExternalDashboards();
  } catch(ee) {
    
  }
}

function cleanPushDashboards() {
  //return;
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD).getSheetByName(DASHBOARD_URLS_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var MASTER_URL = 'https://docs.google.com/spreadsheets/d/1gFB9css5eHo-OKup8jOyWQXtUNe9sgdH98bjhVnV-qQ/edit#gid=1988141836';
  for(var k in data) {
    //if(data[k][0] != 'Romania') { continue; }
    
    var names = {};
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "'+data[k][0]+'"').get();
    while(accounts.hasNext()) {
      var acc = accounts.next()
      var accName = acc.getName();
      names[accName] = acc.getCustomerId();
    }
    
    //Logger.log(Object.keys(names));
    try {
      cleanDashboard(data[k][1], names, MASTER_URL, 'Master');
      cleanDashboard(data[k][1], names, MASTER_URL, 'Analytics');      
      cleanDashboard(data[k][1], names, MASTER_URL, 'Reports');   
      cleanReports(data[k][1], names);
    } catch(ex) {
      Logger.log(data[k][0] + ' : ' + ex); 
    }
  }
}

function cleanFranchiseeDashboards() {
  //return;
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD).getSheetByName('Franchisee Dashboard Urls').getDataRange().getValues();
  data.shift();

  var MASTER_URL = 'https://docs.google.com/spreadsheets/d/1KKyOuyxSN-FlpUt4xyRyevIHhWBEOlZopTpMFsPT_74/edit';
  for(var k in data) {
    //if(data[k][0] != 'Romania') { continue; }
    
    var names = {};
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "'+data[k][0]+'"').get();
    while(accounts.hasNext()) {
      var acc = accounts.next()
      var accName = acc.getName();
      names[accName] = acc.getCustomerId();
    }
    
    //Logger.log(Object.keys(names));
    try {
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 1');
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 2');      
      cleanReports(data[k][1], names);
    } catch(ex) {
      Logger.log(data[k][0] + ' : ' + ex); 
    }
  }
}

function cleanExternalDashboards() {
  //return;
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD).getSheetByName('External Dashboard Urls').getDataRange().getValues();
  data.shift();

  var MASTER_URL = 'https://docs.google.com/spreadsheets/d/1Zs0DT7-y5xezWGOHS-eh9myvw-nQQnsFs0ITZpQz26I/edit#gid=632301538';
  for(var k in data) {
    //if(data[k][0] != 'Romania') { continue; }
    
    var names = {};
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "'+data[k][0]+'"').get();
    while(accounts.hasNext()) {
      var acc = accounts.next()
      var accName = acc.getName();
      names[accName] = acc.getCustomerId();
    }
    
    //Logger.log(Object.keys(names));
    try {
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 1');
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 2');
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 3');
      cleanDashboard(data[k][1], names, MASTER_URL, 'Set 4');      
      cleanReports(data[k][1], names);
    } catch(ex) {
      Logger.log(data[k][0] + ' : ' + ex); 
    }
  }
}

function cleanReports(URL,names) {
  var ss = SpreadsheetApp.openByUrl(URL);
  var configSheet = ss.getSheetByName('Report Links');
  if(!configSheet) { return; }
  
  var data = configSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  
  for(var x in data) {
    if(!data[x][3] || !data[x][1]) { continue; }
    cleanReport(data[x][1], names);
  }
}

function cleanReport(URL, names) {
  var ss = SpreadsheetApp.openByUrl(URL);
  //ss.rename(ss.getName().replace('Elliot', 'Romania'));
  var sheets = ss.getSheets();
  
  var MAP = {
    'Account Name': 1, 'Template': 1, 'Device Report': 1, 'Network Report': 1, 'AdGroup Audit': 1
  }
  
  for(var x in sheets) {
    if(sheets[x].isSheetHidden()) { continue; }
    if(MAP[sheets[x].getName()]) { continue; }
       
    if(names[sheets[x].getName()]) { continue; }
    
    //Logger.log('Deleting ' + sheets[x].getName());
    try {
      ss.deleteSheet(sheets[x]);
    } catch(ex) {
      Logger.log(URL + " :: " +  sheets[x].getName());
    }
  }
}

function cleanDashboard(URL,names,MASTER_URL,configTabName) {
  var ss = SpreadsheetApp.openByUrl(URL);
  var configSheet = SpreadsheetApp.openByUrl(MASTER_URL).getSheetByName(configTabName);
  if(!configSheet) { return; }
  var configSheetData = configSheet.getDataRange().getValues();
  configSheetData.shift();
  configSheetData.shift();
  
  for(var k in configSheetData) {
    if(!configSheetData[k][4]) { continue; }
    TAB_MAP[configSheetData[k][4]] = 1;
  }
  
  var sheets = ss.getSheets();
  for(var j in sheets) {
    var sheetName = sheets[j].getName();
    if(!TAB_MAP[sheetName]) { continue; } 
    var existingNames = [];
    var inputData = sheets[j].getDataRange().getValues();
    inputData.shift();
    inputData.shift();
    
    var toDelete = [];
    for(var l in inputData) {
      existingNames.push(inputData[l][0]);
      if(!names[inputData[l][0]]) {
        toDelete.push(parseInt(l,10)+3); 
        continue;
      }
    }
    
    var newNames = [];
    for(var name in names) {
      if(existingNames.indexOf(name) < 0) {
        if(sheetName == "Account Inputs") { 
          newNames.push([name, names[name]]); 
        } else {
          newNames.push([name]);           
        }
      }
    }
    
    if(newNames.length > 0) {
      sheets[j].getRange(sheets[j].getLastRow()+1,1,newNames.length,newNames[0].length).setValues(newNames);
    }
    
    try {
      for(var l=toDelete.length-1; l >= 0; l--) {
        sheets[j].deleteRow(toDelete[l]);
      }
    } catch(ex) {
      Logger.log([LABEL,URL,sheets[j].getName(),ex].join(' : ')); 
    }
  } 
}