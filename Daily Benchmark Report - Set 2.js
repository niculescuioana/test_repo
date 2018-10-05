/******************************************
* Report - Daily Benchmark Report
* @version: 5.0
* @author: Naman Jindal (nj.itprof@gmail.com)
******************************************/

var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';
var TAB_NAME = 'Benchmark Urls';

function main() {
  
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var hour = NOW.getHours();
  
  var index = hour%2;

  var SETTINGS_MAP = readInputsForAccounts();
  
  Logger.log('Running for: ' + data[index][0]);
  cleanSheets(data[index][0], data[index][1]);
  
  var accountIter = MccApp.accounts()
  .withCondition('LabelNames CONTAINS "'+data[index][0]+'"')
  .orderBy('Impressions ASC')
  .forDateRange('LAST_30_DAYS')
  .get();
  
  while(accountIter.hasNext()) { 
    var account = accountIter.next();
    MccApp.select(account);
    if(!AdWordsApp.currentAccount().getName()) { continue; }    
    
    var SETTINGS = SETTINGS_MAP[AdWordsApp.currentAccount().getName()];
    if(!SETTINGS) {
      SETTINGS = {};
    }
    
    SETTINGS.LABEL = data[index][0];
    SETTINGS.REPORT_URL = data[index][1];
    runScript(SETTINGS);
  }
}




function cleanSheets(LABEL, URL) {
  var accounts = ['Template'];
  var accountIter = MccApp.accounts().withCondition('LabelNames CONTAINS "'+LABEL+'"').withLimit(50).get();
  while(accountIter.hasNext()) {
    var account = accountIter.next();
    accounts.push(account.getName());
  }
  
  var ss = SpreadsheetApp.openByUrl(URL);
  var sheets = ss.getSheets();
  for(var k in sheets) {
    if(accounts.indexOf(sheets[k].getName()) > -1) { continue; }
    ss.deleteSheet(sheets[k])
  }
}

function runScript(SETTINGS) {
  log('Started');
  
  var ss = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL);
  
  var dateKeys = {};
  var accName = AdWordsApp.currentAccount().getName();
  var budget = SETTINGS.DAILY_BUDGET ? SETTINGS.DAILY_BUDGET : '';
  var targetCpa = SETTINGS.CPA_TARGET ? SETTINGS.CPA_TARGET : '';
  
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['Date','Impressions','Clicks','Conversions','Cost','AverageCpc',
              'AveragePosition','CostPerConversion','ConversionValue','Ctr','ConversionRate'];
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'during','LAST_14_DAYS'].join(' ');
  
  var stats = { clicks: [], cpc: [], conversions: [], spends: [], cpa: [], rtos: [], pos: [], ctr: [], cr: [] }
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    dateKeys[row.Date] = 1;
    //row.Date = row.Date;
    stats.clicks.push([row.Date, row.Clicks]);
    stats.conversions.push([row.Date, row.Conversions]);
    stats.cpc.push([row.Date, row.AverageCpc]);
    stats.spends.push([row.Date, row.Cost, budget]);
    stats.cpa.push([row.Date, row.CostPerConversion, targetCpa]);
    stats.pos.push([row.Date, row.AveragePosition]);
    stats.ctr.push([row.Date, row.Ctr]);
    stats.cr.push([row.Date, row.ConversionRate]);
    
    var cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    var rtos = cost == 0 ? 0 : parseFloat(row.ConversionValue.toString().replace(/,/g,'')) / cost; 
    stats.rtos.push([row.Date, rtos]);
  }
  
  var sheet = ss.getSheetByName(AdWordsApp.currentAccount().getName());
  if(!sheet) {
    var templateSheet = ss.getSheetByName('Template');
    ss.setActiveSheet(templateSheet);
    sheet = ss.duplicateActiveSheet();    
    sheet.setName(AdWordsApp.currentAccount().getName());
  }
  
  sheet.showSheet();
  //ss.setActiveSheet(sheet);
  //ss.moveActiveSheet(SETTINGS.pos);
  
  sheet.getRange(3,2,stats.clicks.length,2).setValues(stats.clicks).sort({column: 2, ascending: true});
  sheet.getRange(20,2,stats.spends.length,3).setValues(stats.spends).sort({column: 2, ascending: true});
  sheet.getRange(37,2,stats.cpc.length,2).setValues(stats.cpc).sort({column: 2, ascending: true});
  sheet.getRange(54,2,stats.conversions.length,2).setValues(stats.conversions).sort({column: 2, ascending: true});
  sheet.getRange(71,2,stats.cpa.length,3).setValues(stats.cpa).sort({column: 2, ascending: true});
  sheet.getRange(88,2,stats.pos.length,2).setValues(stats.pos).sort({column: 2, ascending: true});  
  sheet.getRange(105,2,stats.ctr.length,2).setValues(stats.ctr).sort({column: 2, ascending: true});  
  sheet.getRange(122,2,stats.cr.length,2).setValues(stats.cr).sort({column: 2, ascending: true});  
  
  
  log('Finished');
}

function readInputsForAccounts() {
  Logger.log('Reading Budget & Target CPA');
  
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


function log(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);  
}


function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}