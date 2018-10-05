/*************************************************
* Analytics Report
* @version: 3.0
* @author: Naman Jindal (nj.itprof@gmail.com)
***************************************************/

var MASTER_DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1643684505';
var TAB_NAME = 'Analytics - All Profiles';

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1H3onoG-Pwi6f1GXWqyD2d2MOw3W7E9h87VYuF9Lm-FI/edit#gid=1179113750';
var REPORT_TAB = 'Analytics Report';


function main() {
  
  cleanupSheet();
  
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'),10);
  //Logger.log(hour);
  //return;
  if(hour == 0) {
    exportAnalyticsAccounts();
  }
  
  var reportHours = [7,10,15,18,21];
  
  if(reportHours.indexOf(hour) < 0) {
    return;
  }
  
  var ROW_BY_UA_NUM = {};
  var outputSheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(REPORT_TAB);
  
  var data = outputSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var TM = getAdWordsFormattedDate(0, 'MMM yyyy');
  var LM = getLastDay('MMM yyyy');
  var PM = getPriorMonthDay('MMM yyyy');
  
  var TMLY = TM.split(' ')[0] + ' ' + (parseInt(TM.split(' ')[1],10)-1);
  var LMLY = LM.split(' ')[0] + ' ' + (parseInt(LM.split(' ')[1],10)-1);
  var PMLY = PM.split(' ')[0] + ' ' + (parseInt(PM.split(' ')[1],10)-1);
  
  for(var j in data) {
    if(!data[j][1]) { continue; }
    //if(j != 136) { continue; }
    var SETTINGS = {};
    SETTINGS.TM = TM;
    SETTINGS.LM = LM;
    SETTINGS.PM = PM;
    
    SETTINGS.TMLY = TMLY;
    SETTINGS.LMLY = LMLY;            
    SETTINGS.PMLY = PMLY;            
    
    SETTINGS.outputSheet = outputSheet;
    SETTINGS.rNum = parseInt(j)+3;
    SETTINGS.UA_NUMBER = data[j][1];
    SETTINGS.ACC_NAME = data[j][0];
    if(!SETTINGS.UA_NUMBER) { continue; }
    
    try {
      exportDataForAccount(SETTINGS,ROW_BY_UA_NUM);
    } catch(ex) {
      Logger.log(ex);
    }
    
    //Logger.log(SETTINGS.rNum);
    //break;
    if(shouldExitNow()) { Logger.log('Running out of time'); break; }
  }
  
  var headerRow = [getAdWordsFormattedDate(2, 'MMM dd, yyyy'), getAdWordsFormattedDate(3, 'MMM dd, yyyy'), '30 Day Avg',
                   TM, LM, PM, TMLY, LMLY, PMLY];
  
  outputSheet.getRange(2,5,1,headerRow.length).setValues([headerRow]);
  outputSheet.getRange(2,16,1,headerRow.length).setValues([headerRow]);
  outputSheet.getRange(2,27,1,1).setValue(getAdWordsFormattedDate(2, 'MMM dd, yyyy'));
  
  //outputSheet.getRange(3, 3, outputSheet.getLastRow(), 9).setNumberFormat("#,##0.00"); // Revenue
  //outputSheet.getRange(3, 12, outputSheet.getLastRow(), 9).setNumberFormat("#,##0"); // Trasnsactions  
  
  formatReport(outputSheet);
}  

function formatReport(outputSheet) {
  
  var data = outputSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var backgrounds = [];
  for(var k in data) {
    var avgRevenue = data[k][6];
    var min = avgRevenue*0.95;
    var max = avgRevenue*1.05;
    
    var colors = [];
    var yestRevenue = data[k][3];
    
    if(yestRevenue < min) {
      colors.push('#f4cccc'); //red
    } else if(yestRevenue > max) {
      colors.push('#d9ead3'); //green
    } else {
      colors.push('#fff2cc'); //yellow
    } 
    
    if(data[k][4] < min) {
      colors.push('#f4cccc'); //red
    } else if(data[k][4] > max) {
      colors.push('#d9ead3'); //green
    } else {
      colors.push('#fff2cc'); //yellow
    } 
    
    if(data[k][5] < min) {
      colors.push('#f4cccc'); //red
    } else if(data[k][5] > max) {
      colors.push('#d9ead3'); //green
    } else {
      colors.push('#fff2cc'); //yellow
    } 
    backgrounds.push(colors);
  }  
  
  outputSheet.getRange(3,4,backgrounds.length,backgrounds[0].length).setBackgrounds(backgrounds);
}

function shouldExitNow() {
  return (AdWordsApp.getExecutionInfo().getRemainingTime() < 60*3) 
}

function exportDataForAccount(SETTINGS, ROW_BY_UA_NUM) {
  var ACCOUNT_TIME = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  
  var row = ROW_BY_UA_NUM[SETTINGS.UA_NUMBER];
  if(!row) {
    row = gatherStats(SETTINGS);
    ROW_BY_UA_NUM[SETTINGS.UA_NUMBER] = row;
  }
  
  row[0] = SETTINGS.ACC_NAME;
  row[1] = SETTINGS.UA_NUMBER;
  
  if(row.length == 0) { return; }
  
  SETTINGS.outputSheet.getRange(SETTINGS.rNum, 1, 1, row.length).setValues([row]);
}




function gatherStats(SETTINGS) {
  var ids = getLinkedProfileIds(SETTINGS.UA_NUMBER);
  if(ids.length == 0) { return []; }
  
  var FROM = getAdWordsFormattedDate(30, 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var today = getAdWordsFormattedDate(0, 'yyyyMMdd');
  var stats = initStatsMap(SETTINGS,today);
  
  var results = {};
  for(var k in ids) {
    ids[k] = ids[k].toString().trim();
    results[ids[k]] =  getAnalyticsStatsByDate(ids[k],FROM,TO,stats,today);
  }
  
  for(var key in results) {
    for(var dateKey in stats) {
      if(!results[key][dateKey]) { continue; }
      stats[dateKey].transactions += results[key][dateKey].Transactions;
      stats[dateKey].revenue += results[key][dateKey].Revenue;
    }
  }
  
  //Logger.log(stats['Total']);
  stats['Total'].transactions = Math.round(stats['Total'].transactions/30);
  stats['Total'].revenue = (stats['Total'].revenue/30).toFixed(2);
  
  //Get Stats For This Month
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  //Get Stats For Last Month
  var DT = getLastDay('yyyy-MM-dd');
  var year = parseInt(DT.substring(0,4),10) - 1;
  DT = year + '-' + DT.substring(5,10);
  
  var FROM = [DT.split('-')[0],DT.split('-')[1]].join('-') + '-' + '01';
  for(var k in ids) {
    getAnalyticsStats(ids[k],FROM,TO,stats);
  }
  
  var convMap = initConvMap();
  var results = {};
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var FROM = getAdWordsFormattedDate(2, 'yyyy-MM-dd');  
  for(var k in ids) {
    results[ids[k]] = getCpcSourceGoalCompletions(ids[k],FROM,TO);
  }
  
  for(var key in results) {
    for(var dateKey in convMap) {
      if(!results[key][dateKey]) { continue; }
      convMap[dateKey].conv += results[key][dateKey].conv;
    }
  }
  
  var outputRow = [SETTINGS.ACC_NAME,SETTINGS.UA_NUMBER];
  for(var key in stats) {
    outputRow.push(stats[key].revenue);
  }
  
  for(var key in stats) {
    outputRow.push(stats[key].transactions);
  }
  
  for(var key in convMap) {
    outputRow.push(convMap[key].conv);
  }
  
  //Logger.log(outputRow);
  return outputRow;
}

function getLinkedProfileIds(UA_NUMBER) {

  return UA_NUMBER.toString().split(',');
  //}
  
  return '';
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var profiles = [];
  for(var k in data) {
    if(UA_NUMBER.indexOf(data[k][2]) > -1) {
      profiles.push(data[k][3]);
    }
  }
  
  return profiles;
}

function initConvMap() {
  var map = {};
  
  var today = getAdWordsFormattedDate(0, 'yyyyMMdd');
  map[today] = { conv:0 }
  
  var yesterday = getAdWordsFormattedDate(1, 'yyyyMMdd');
  map[yesterday] = { conv:0 }
  
  var twoDaysBack = getAdWordsFormattedDate(2, 'yyyyMMdd');
  map[twoDaysBack] = { conv:0 }
  
  return map;  
}

function initStatsMap(SETTINGS,today) {
  var map = {};
  
  map[today] = { transactions:0, revenue:0 }
  
  var yesterday = getAdWordsFormattedDate(1, 'yyyyMMdd');
  map[yesterday] = { transactions:0, revenue:0 }
  
  var twoDaysBack = getAdWordsFormattedDate(2, 'yyyyMMdd');
  map[twoDaysBack] = { transactions:0, revenue:0 }
  
  var threeDaysBack = getAdWordsFormattedDate(3, 'yyyyMMdd');
  map[threeDaysBack] = { transactions:0, revenue:0 }
  
  map['Total'] = { transactions:0, revenue:0 }
  map[SETTINGS['TM']] = { transactions:0, revenue:0 }
  map[SETTINGS['LM']] = { transactions:0, revenue:0 }
  map[SETTINGS['PM']] = { transactions:0, revenue:0 }
  
  map[SETTINGS['TMLY']] = { transactions:0, revenue:0 }
  map[SETTINGS['LMLY']] = { transactions:0, revenue:0 }
  map[SETTINGS['PMLY']] = { transactions:0, revenue:0 }
  
  return map;
}

function getCpcSourceGoalCompletions(id,FROM,TO) {
  var optArgs = { 'dimensions': 'ga:date', 'filters': 'ga:medium==cpc' };
  var attempts = 3;
  var results;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      results = Analytics.Data.Ga.get(
        'ga:'+id,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:goalCompletionsAll",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = results.getRows();
  
  var reportRows = {};
  for(var k in rows) {
    reportRows[rows[k][0]] = { conv: parseInt(rows[k][1],10) }
  }
  
  return reportRows;
}

function getAnalyticsStatsByDate(id,FROM,TO,stats,today) {
  var optArgs = { 'dimensions': 'ga:date' };
  var attempts = 3;
  var results;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      // Make a request to the API.
      results = Analytics.Data.Ga.get(
        'ga:'+id,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = results.getRows();
  
  var reportRows = {};
  for(var k in rows) {
    reportRows[rows[k][0]] = { Transactions: parseInt(rows[k][1],10),  Revenue: parseFloat(rows[k][2]) }
    if(rows[k][0] == today) { continue; }
    stats['Total'].transactions += reportRows[rows[k][0]].Transactions;
    stats['Total'].revenue += reportRows[rows[k][0]].Revenue;
  }
  
  return reportRows;
}

function getAnalyticsStats(id,FROM,TO,stats) {
  var attempts = 3;
  var results;
  
  var optArgs = { 'dimensions': 'ga:yearMonth' }
  // Make a request to the API.
  while(attempts > 0) {
    try {
      results = Analytics.Data.Ga.get(
        'ga:'+id,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = results.getRows();
  
  var monthNames = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  for(var k in rows) {
    var key = monthNames[parseInt(rows[k][0].toString().substring(4,6),10)] + ' ' + rows[k][0].toString().substring(0,4);
    if(!stats[key]) { continue; }
    stats[key].transactions += parseInt(rows[k][1],10);
    stats[key].revenue += parseFloat(rows[k][2]);
  }
}


function getLastDay(format) {
  var date = new Date();
  date.setDate(0);  
  date.setHours(10);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function getPriorMonthDay(format) {
  var date = new Date();
  date.setDate(0);
  date.setDate(0);  
  date.setHours(10);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
} 



/************************** Analytic's Profile Refresh ************************************/
function exportAnalyticsAccounts() {
  var rows = getProfile();
  SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL).getSheetByName(TAB_NAME).getRange(1,1,rows.length,rows[0].length).setValues(rows);  
}

function getProfile() {
  var accounts = Analytics.Management.Accounts.list();
  var rows = [['AccountId', 'AccountName', 'Web Property Id', 'Profile Id']];
  if (accounts.getItems()) {
    for(var i in accounts.getItems()){
      var accountId = accounts.getItems()[i].getId();
      var accountName = accounts.getItems()[i].getName();  
      var webProperties = Analytics.Management.Webproperties.list(accountId);
      if (webProperties.getItems()) {
        for(var j in webProperties.getItems()) {
          var webPropertyId = webProperties.getItems()[j].getId();
          var profiles;
          try {
            profiles = Analytics.Management.Profiles.list(accountId, webPropertyId);          
          } catch(ex) {
            Utilities.sleep(1000);
            profiles = Analytics.Management.Profiles.list(accountId, webPropertyId);          
          }
          if (profiles.getItems()) {
            for(var k in profiles.getItems()) {
              var profile = profiles.getItems()[k];
              rows.push([accountId,accountName,webPropertyId,profile.getId()]);
            }	
          } else {
            rows.push([accountId,accountName,webPropertyId,'']);
          }
        }
      } else {
        rows.push([accountId,accountName,'','']);
      }
      Utilities.sleep(1500);
    }
  } else {
    throw new Error('No accounts found.');
  }
  
  return rows;
}




function readInputsForAccounts() {
  var SETTINGS = {};
  
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName('Dashboard Urls').getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][1]) { continue; }
    var sheet = SpreadsheetApp.openByUrl(data[k][1]).getSheetByName('Account Inputs');
    if(!sheet) {
      continue;
    }
    
    var inputData = sheet.getDataRange().getValues();
    inputData.shift(); 
    var header = inputData.shift()
    
    for(var j in inputData) {
      SETTINGS[inputData[j][0]] = {};
      for(var l in header) {
        SETTINGS[inputData[j][0]][header[l]] = inputData[j][l];
      }
    }
  }
  
  return SETTINGS;
}

function cleanupSheet() {
  var SETTINGS = readInputsForAccounts();
  var outputSheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(REPORT_TAB);
  
  var data = outputSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var toRemove = [], map = {};
  for(var z in data) {
    if(!SETTINGS[data[z][0]]) {
      toRemove.push(parseInt(z,10)+3); 
      continue;
    }
    
    /*if(map[data[z][0]]) {
      toRemove.push(parseInt(z,10)+3); 
      continue;
    }*/
    
    if(SETTINGS[data[z][0]]['PROFILE_IDS']) {
      data[z][1] = SETTINGS[data[z][0]]['PROFILE_IDS'];
    }
    
    delete SETTINGS[data[z][0]];
  }
  
  outputSheet.getRange(3,1,data.length,data[0].length).setValues(data);
  
  var z = toRemove.length - 1;
  for(;z>=0;z--) {
    outputSheet.deleteRow(toRemove[z]);
  }
  
  
  
  var out = [];
  for(var name in SETTINGS) {
    if(SETTINGS[name]['PROFILE_IDS']) {
      out.push([name, SETTINGS[name]['PROFILE_IDS']]);
    }
  }
  
  if(out.length) {
    outputSheet.getRange(outputSheet.getLastRow()+1, 1, out.length, out[0].length).setValues(out);
  }
}