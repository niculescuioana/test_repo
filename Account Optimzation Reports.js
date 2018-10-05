// To Update
var RESULTS_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1EcdZ17IUtsztXGGoGfAvOdg7amggMJ8I8Cqdbu6Q7lA/edit?ts=5ba120e1#gid=1312575355';


var RESPONSE_URL = 'https://docs.google.com/spreadsheets/d/1T9tiaAThIkuDSGs5UpwsJqmMAUzvrVRevJrS_eAUtwg/edit#gid=1544056572';
var RESPONSE_TAB_NAME = '% Done';

function main() {
  
  //compileResponseReport();
  compileAccountOptimizationReports();
  compileAccountOptimizationReportsOthers();
  
}


function compileAccountOptimizationReports() {
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var TAB_NAME = 'Dashboard Urls';
  
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME);
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var today = Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy');
  var accountsByManager = {}, reportByManager = {};
  for(var k in data) {
    var manager = data[k][0];
    accountsByManager[manager] = {};
    if(!data[k][4]) {
      var ss = SpreadsheetApp.create('Account Optimization Report - ' + manager + ' ('+today+')');
      data[k][4] = ss.getUrl();
    }
    
    reportByManager[manager] = data[k][4]
    
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "' + manager + '"').get();
    while(accounts.hasNext()) {
      var account = accounts.next();
      accountsByManager[manager][account.getCustomerId().replace(/-/g,'')] = account.getName();
    }
  } 
  
  tab.getRange(2,1,data.length,data[0].length).setValues(data);
  
  var ignoreColumns = {
    'Priority': 1, 'Customer Id': 1, 'Account Name': 1, 'Account Description': 1, 'Mccs': 1, 
    'Partner Name': 1, 'Account Reason': 1, 'Campaign Id': 1, 'Offering Last Updated': 1
  }
  
  var ignoreTabs = {
    'Ad Rotation': 1,
    'Account Improvement Summary': 1,
    'Account Improvement Details': 1
  }
  
  var optimizationsByAccount = {};
  
  var source = SpreadsheetApp.openByUrl(RESULTS_REPORT_URL);
  var tabs = source.getSheets();
  
  for(var x in tabs) {
    if(ignoreTabs[tabs[x].getName()]) { continue; }
    var data = tabs[x].getDataRange().getValues();
    
    var issueType = tabs[x].getName();
    data.shift();
    data.shift();
    data.shift();
    var pitch = data.shift();
    
    var recommededForAll = [], recommendAllFlag = false, prorityFlag = false, reasonIndex, sheetHeader;
    for(var z in data) {
      if(data[z][0] == ' Recommended for these Accounts because they meet ALL of these conditions:') {
        recommendAllFlag = true;
        continue; 
      }
      
      if(recommendAllFlag) {
        if(data[z][0].indexOf('*') < 0) {
          recommendAllFlag = false;
          continue;
        } else {
          recommededForAll.push(data[z][0]); 
        }
      }
      
      if(data[z][0] == 'Priority') {
        prorityFlag = true;
        sheetHeader = data[z];
        reasonIndex = data[z].indexOf('Account Reason');
        continue;
      }
      
      if(!prorityFlag) {
        continue;
      }
      
      if(!optimizationsByAccount[data[z][1]]) {
        optimizationsByAccount[data[z][1]] = [];
      }
      
      var accountReasons = [], campaignName = '';
      if(reasonIndex != -1 && data[z][reasonIndex]) {
        accountReasons = recommededForAll.concat(data[z][reasonIndex]);
      } else {
        accountReasons =  recommededForAll;
      }
      
      var details = [];
      for(var n in sheetHeader) {
        if(ignoreColumns[sheetHeader[n]] || !sheetHeader[n] || !data[z][n]) { continue; }
        details.push([sheetHeader[n], data[z][n]].join(': '));
      }
      
      var row = [data[z][0], issueType, pitch, details.join('\n'), accountReasons.join('\n'), '', ''];
      optimizationsByAccount[data[z][1]].push(row);
    }
  }
  
  var header = [['Priority', 'Type', 'Elevator Pitch', 'Details', 'Reasons', 'Action Taken?', 'Comments']];
  for(var manager in accountsByManager) {
    var tabMap = {};
    var ss = SpreadsheetApp.openByUrl(reportByManager[manager]);
    for(var id in accountsByManager[manager]) {
      if(!optimizationsByAccount[id]) { continue; }
      var accountName = accountsByManager[manager][id];
      
      var tabName = accountName+':'+id;
      tabMap[tabName] = 1;
      var tab = ss.getSheetByName(tabName);
      if(!tab) { 
        tab = ss.insertSheet(tabName); 
        tab.setFrozenRows(1);
        
        tab.setColumnWidth(1, 60);
        tab.setColumnWidth(2, 185);
        tab.setColumnWidth(3, 300);
        tab.setColumnWidth(4, 300);
        tab.setColumnWidth(5, 275);
        tab.setColumnWidth(6, 120);
        tab.setRowHeight(1, 40);
      }
      
      var rows = header.concat(optimizationsByAccount[id]);
      
      tab.clearContents();
      tab.getRange(1,1,rows.length,rows[0].length).setValues(rows).setFontFamily('Calibri');
      
      tab.getDataRange().setWrap(true).setVerticalAlignment('top');
      tab.getRange(1,1,1,rows[0].length).setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
      
      
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(['y','n']).build();
      tab.getRange(2,6,rows.length-1,1).setDataValidation(rule); 
      //tab.getRange(2,7,rows.length-1,1).clearDataValidations();
      
      if((tab.getMaxColumns() - tab.getLastColumn()) > 0) {
        tab.deleteColumns(tab.getLastColumn()+1, tab.getMaxColumns() - tab.getLastColumn());
      }
    }
    
    var toDel = [], sheets = ss.getSheets();
    for(var z in sheets) {
      if(!tabMap[sheets[z].getName()]) {
        toDel.push(sheets[z]); 
      }
    }
    
    for(var z in toDel) {
      try {
        ss.deleteSheet(toDel[z]);
      } catch(ex) {
        Logger.log(manager);
      }
    }
  }
}


function compileAccountOptimizationReportsOthers() {
  var today = Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy');
  var URL = 'https://docs.google.com/spreadsheets/d/1sC8SBAElHloEO3JmNuuWwDCs3DFZWbtFj6lLYrYYshk/edit#gid=1427124378';
  //var URL = SpreadsheetApp.create('Account Optimization Report - Others ('+today+')').getUrl();
  //Logger.log(URL);
  
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var TAB_NAME = 'Dashboard Urls';
  
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME);
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var accounts = MccApp.accounts();
  
  var manager = 'Others';
  
  var accountsByManager = {};
  accountsByManager[manager] = {};
  for(var k in data) {
    accounts = accounts.withCondition('LabelNames DOES_NOT_CONTAIN "' + data[k][0] + '"');
  }
  
  accounts = accounts.get();
  while(accounts.hasNext()) {
    var account = accounts.next();
    accountsByManager[manager][account.getCustomerId().replace(/-/g,'')] = account.getName();
  }
  
  
  var ignoreColumns = {
    'Priority': 1, 'Customer Id': 1, 'Account Name': 1, 'Account Description': 1, 'Mccs': 1, 
    'Partner Name': 1, 'Account Reason': 1, 'Campaign Id': 1, 'Offering Last Updated': 1
  }
  
  var ignoreTabs = {
    'Ad Rotation': 1,
    'Account Improvement Summary': 1,
    'Account Improvement Details': 1
  }
  
  var optimizationsByAccount = {};
  var source = SpreadsheetApp.openByUrl(RESULTS_REPORT_URL);
  var tabs = source.getSheets();
  
  for(var x in tabs) {
    if(ignoreTabs[tabs[x].getName()]) { continue; }
    var data = tabs[x].getDataRange().getValues();
    
    var issueType = tabs[x].getName();
    data.shift();
    data.shift();
    data.shift();
    var pitch = data.shift();
    
    var recommededForAll = [], recommendAllFlag = false, prorityFlag = false, reasonIndex, sheetHeader;
    for(var z in data) {
      if(data[z][0] == ' Recommended for these Accounts because they meet ALL of these conditions:') {
        recommendAllFlag = true;
        continue; 
      }
      
      if(recommendAllFlag) {
        if(data[z][0].indexOf('*') < 0) {
          recommendAllFlag = false;
          continue;
        } else {
          recommededForAll.push(data[z][0]); 
        }
      }
      
      if(data[z][0] == 'Priority') {
        prorityFlag = true;
        sheetHeader = data[z];
        reasonIndex = data[z].indexOf('Account Reason');
        continue;
      }
      
      if(!prorityFlag) {
        continue;
      }
      
      if(!optimizationsByAccount[data[z][1]]) {
        optimizationsByAccount[data[z][1]] = [];
      }
      
      var accountReasons = [], campaignName = '';
      if(reasonIndex != -1 && data[z][reasonIndex]) {
        accountReasons = recommededForAll.concat(data[z][reasonIndex]);
      } else {
        accountReasons =  recommededForAll;
      }
      
      var details = [];
      for(var n in sheetHeader) {
        if(ignoreColumns[sheetHeader[n]] || !sheetHeader[n] || !data[z][n]) { continue; }
        details.push([sheetHeader[n], data[z][n]].join(': '));
      }
      
      var row = [data[z][0], issueType, pitch, details.join('\n'), accountReasons.join('\n'), '', ''];
      optimizationsByAccount[data[z][1]].push(row);
    }
  }
  
  var header = [['Priority', 'Type', 'Elevator Pitch', 'Details', 'Reasons', 'Action Taken?', 'Comments']];
  for(var manager in accountsByManager) {
    var tabMap = {};
    var ss = SpreadsheetApp.openByUrl(URL);
    for(var id in accountsByManager[manager]) {
      if(!optimizationsByAccount[id]) { continue; }
      var accountName = accountsByManager[manager][id];
      
      var tabName = accountName+':'+id;
      tabMap[tabName] = 1;
      var tab = ss.getSheetByName(tabName);
      if(!tab) { 
        tab = ss.insertSheet(tabName); 
        tab.setFrozenRows(1);
        
        tab.setColumnWidth(1, 60);
        tab.setColumnWidth(2, 185);
        tab.setColumnWidth(3, 300);
        tab.setColumnWidth(4, 300);
        tab.setColumnWidth(5, 275);
        tab.setColumnWidth(6, 120);
        tab.setRowHeight(1, 40);
      }
      
      var rows = header.concat(optimizationsByAccount[id]);
      
      tab.clearContents();
      tab.getRange(1,1,rows.length,rows[0].length).setValues(rows).setFontFamily('Calibri');
      
      tab.getDataRange().setWrap(true).setVerticalAlignment('top');
      tab.getRange(1,1,1,rows[0].length).setBackground('#efefef').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
      
      
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(['y','n']).build();
      tab.getRange(2,6,rows.length-1,1).setDataValidation(rule); 
      //tab.getRange(2,7,rows.length-1,1).clearDataValidations();
      
      if((tab.getMaxColumns() - tab.getLastColumn()) > 0) {
        tab.deleteColumns(tab.getLastColumn()+1, tab.getMaxColumns() - tab.getLastColumn());
      }
    }
    
    var toDel = [], sheets = ss.getSheets();
    for(var z in sheets) {
      if(!tabMap[sheets[z].getName()]) {
        toDel.push(sheets[z]); 
      }
    }
    
    for(var z in toDel) {
      ss.deleteSheet(toDel[z]);
    }
  }
}

function compileResponseReport() {
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var TAB_NAME = 'CS1 Reports';
  
  var initMap = {
    'Isuru': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Yannis': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Monique': { 'Blanks': 0, 'Yes': 0, 'No': 0 },    
    'Keerat': { 'Blanks': 0, 'Yes': 0, 'No': 0 },    
    'Neeraj': { 'Blanks': 0, 'Yes': 0, 'No': 0 },        
    'Ian': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Sandeep': { 'Blanks': 0, 'Yes': 0, 'No': 0 },            
    'Elliot': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Leslie': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Reece': { 'Blanks': 0, 'Yes': 0, 'No': 0 },    
    'Tariq': { 'Blanks': 0, 'Yes': 0, 'No': 0 },
    'Other': { 'Blanks': 0, 'Yes': 0, 'No': 0 }
  }
   
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME);
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var resultsByManager = {}, overview = JSON.parse(JSON.stringify(initMap)), resultsByType = {};
  var today = Utilities.formatDate(new Date(), 'GMT', 'd MMMM');
  for(var k in data) {
    var manager = data[k][0];
    if(!initMap[manager]) { continue; }
    var url = data[k][2];
    //resultsByManager[manager] = {};
    
    var tabs = SpreadsheetApp.openByUrl(url).getSheets();
    for(var z in tabs) {
      var inData = tabs[z].getDataRange().getValues();
      inData.shift();
      
      for(var i in inData) {
        if(!inData[i][1]) { continue; }
        var type = inData[i][1];
        if(!resultsByManager[type]) {
          resultsByManager[type] = JSON.parse(JSON.stringify(initMap));
        }
        
        if(!resultsByType[type]) {
          resultsByType[type] = { 'Blanks': 0, 'Yes': 0, 'No': 0 };
        }
        
        if(inData[i][5] == 'n') {
          overview[manager]['No']++;
          resultsByManager[type][manager]['No']++;
          resultsByType[type]['No']++;
        } else if(inData[i][5] == 'y') { 
          overview[manager]['Yes']++;          
          resultsByManager[type][manager]['Yes']++;          
          resultsByType[type]['Yes']++;          
        } else {
          overview[manager]['Blanks']++;          
          resultsByManager[type][manager]['Blanks']++; 
          resultsByType[type]['Blanks']++;          
        }
      }
    }
  }   
  
  
  var tab = SpreadsheetApp.openByUrl(RESPONSE_URL).getSheetByName(RESPONSE_TAB_NAME);
  tab.clearContents();
  
  var output = [['','','Isuru','Yannis','Monique','Keerat','Neeraj',        
                 'Ian','Sandeep','Elliot','Leslie','Reece','Tariq','Other']];
  
  for(var type in resultsByManager) {
    var rows = [[type, 'Blanks'], ['', 'Yes'], ['', 'No'], ['', '']];
    for(var manager in  resultsByManager[type]) {
      var total = resultsByManager[type][manager]['Blanks'] + resultsByManager[type][manager]['Yes'] + resultsByManager[type][manager]['No'];
      if(total == 0) {
        rows[0].push(0);
        rows[1].push(0);
        rows[2].push(0);        
        rows[3].push('');
      } else {
        rows[0].push(resultsByManager[type][manager]['Blanks']/total);
        rows[1].push(resultsByManager[type][manager]['Yes']/total);
        rows[2].push(resultsByManager[type][manager]['No']/total);        
        rows[3].push('');
      }
        
    }
    
    output = output.concat(rows);
  }
  
  tab.getRange(1,1,output.length,output[0].length).setValues(output);
  
  var tab = SpreadsheetApp.openByUrl(RESPONSE_URL).getSheetByName('Count');
  tab.clearContents();
  
  var output = [['','','Isuru','Yannis','Monique','Keerat','Neeraj',        
                 'Ian','Sandeep','Elliot','Leslie','Reece','Tariq','Other']];
  
  for(var type in resultsByManager) {
    var rows = [[type, 'Blanks'], ['', 'Yes'], ['', 'No'], ['', '']];
    for(var manager in  resultsByManager[type]) {
      rows[0].push(resultsByManager[type][manager]['Blanks']);
      rows[1].push(resultsByManager[type][manager]['Yes']);
      rows[2].push(resultsByManager[type][manager]['No']);        
      rows[3].push('');
    }
    
    output = output.concat(rows);
  }
  
  tab.getRange(1,1,output.length,output[0].length).setValues(output);
  
  var tab = SpreadsheetApp.openByUrl(RESPONSE_URL).getSheetByName('Overview');
  //tab.clearContents();
  
  var output = [['','Isuru','Yannis','Monique','Keerat','Neeraj',        
                 'Ian','Sandeep','Elliot','Leslie','Reece','Tariq','Other']];
  
  var rows = [['Blanks'], ['Yes'], ['No']];
  for(var manager in overview) {
    var total = overview[manager]['Blanks'] + overview[manager]['Yes'] + overview[manager]['No']
    rows[0].push(overview[manager]['Blanks']/total);
    rows[1].push(overview[manager]['Yes']/total);
    rows[2].push(overview[manager]['No']/total);        
  }
  
  output = output.concat(rows);
  tab.getRange(1,1,output.length,output[0].length).setValues(output);
  
  var output = [['', 'Blanks', 'Yes', 'No']];
  for(var type in resultsByType) {
    var total = resultsByType[type]['Blanks'] +  resultsByType[type]['Yes'] + resultsByType[type]['No'];
    output.push([type, resultsByType[type]['Blanks']/total, resultsByType[type]['Yes']/total, resultsByType[type]['No']/total]); 
  }
                 
  tab.getRange(7,1,output.length,output[0].length).setValues(output);
}
