/*************************************************
* Master Daily Management Reporting Script
* @version: 1.0
* @author: Naman Jindal (nj.itprof@gmail.com)
***************************************************/

var MASTER_DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1415541725';
var DASHBOARD_URLS_TAB = 'Dashboard Urls';
var MANEGEMENT_URS_TAB = 'Management Urls';

var MASTER_MANAGEMENT_URL = 'https://docs.google.com/spreadsheets/d/1LVKJ-ti1BcYktSli94Q4RPCzzCyz01fzTk4kYUaO3UI/edit';

var EMAIL_TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1R1d7xalKKRP6SU37cXBHxl-15q7e918sxeWLtpkegJY/edit#gid=733574844&vpid=A1";

var ANALYTICS_STATS_REPORT = 'https://docs.google.com/spreadsheets/d/1H3onoG-Pwi6f1GXWqyD2d2MOw3W7E9h87VYuF9Lm-FI/edit';
var ANALYTICS_TAB_NAME = 'Analytics Report';


function main() {
  
  var index = 1;
  
  var masterSS = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL);
  var data = masterSS.getSheetByName(MANEGEMENT_URS_TAB).getDataRange().getValues();
  data.shift();
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var hour = NOW.getHours();
  
  var ACCOUNT_INPUTS = readAccountSettings();
  
  var LABEL = data[index][0];
  var REPORT_URL = data[index][1];
  
  Logger.log('Running for: ' + LABEL);
  
  var mccAccount = AdWordsApp.currentAccount();
  var accountIter = MccApp.accounts()
  //.withCondition('Name = "Designahousesign.co.uk"')
  .withCondition('LabelNames CONTAINS "'+LABEL+'"')
  .get();
  
  if(accountIter.totalNumEntities() <= 30) {
    MccApp.accounts().withCondition('LabelNames CONTAINS "'+LABEL+'"')
    //.withCondition('Name = "Designahousesign.co.uk"')
    .executeInParallel('runScript','compileResults',JSON.stringify({ 'ACCOUNT_INPUTS': ACCOUNT_INPUTS, 'LABEL': LABEL, 'REPORT_URL': REPORT_URL, 'IS_RUNNING_IN_PARLLEL_MODE': 1 }));
    
    return;
  }
  
  var ids = [];
  while(accountIter.hasNext()) {
    var account = accountIter.next();
    MccApp.select(account);
    
    if(ids.length < 30) {
      ids.push(AdWordsApp.currentAccount().getCustomerId());
      continue;
    }
    
    runScript(JSON.stringify({ 'ACCOUNT_INPUTS': ACCOUNT_INPUTS, 'LABEL': LABEL, 'REPORT_URL': REPORT_URL, 'IS_RUNNING_IN_PARLLEL_MODE': 0 }));
  }
  
  MccApp.select(mccAccount);
  
  MccApp.accounts().withIds(ids).executeInParallel('runScript','compileResults',JSON.stringify({ 'ACCOUNT_INPUTS': ACCOUNT_INPUTS, 'LABEL': LABEL, 'REPORT_URL': REPORT_URL, 'IS_RUNNING_IN_PARLLEL_MODE': 1 }));
}

function runScript(INPUT) {
  if(!AdWordsApp.currentAccount().getName()) { return; }
  //Logger.log(AdWordsApp.currentAccount().getName());
  var INPUT_MAP = JSON.parse(INPUT);
  
  var SETTINGS = INPUT_MAP['ACCOUNT_INPUTS'][AdWordsApp.currentAccount().getName()];
  if(!SETTINGS) {
    SETTINGS = {}
    
    SETTINGS.MANAGER = '';
    SETTINGS.MONTHLY_BUDGET = '';
    SETTINGS.DAILY_BUDGET = '';
    SETTINGS.CPA_TARGET = ''; 
  }
  
  SETTINGS.CLIENT = AdWordsApp.currentAccount().getName();
  
  SETTINGS.REPORT_URL = INPUT_MAP['REPORT_URL'];
  SETTINGS.LABEL = INPUT_MAP['LABEL']; 
  SETTINGS.IS_RUNNING_IN_PARLLEL_MODE = INPUT_MAP['IS_RUNNING_IN_PARLLEL_MODE']; 
  compileDailyReport(SETTINGS);
  compileWeeklyReport(SETTINGS);
}

function compileDailyReport(SETTINGS) {
  if(SETTINGS.LABEL.toLowerCase() == 'new business') {
    compileDailyReportWithKPI(SETTINGS); 
  } else {
    compileDailyReportNoKPI(SETTINGS);  
  }
}

function compileDailyReportWithKPI(SETTINGS) {
  var URLS = [SETTINGS.REPORT_URL];
  
  //if(SETTINGS.FLAG.toUpperCase() != 'Y'){ info('Script turned off for the account. Exiting.'); return; }
  
  
  var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy').toUpperCase();
  
  for(var l in URLS) {      
    SETTINGS.URL = URLS[l];      
    if(SETTINGS.URL.length < 5) {
      info('Spreadsheet Url Not found. Exiting');
      continue;
    } else {
      //info(SETTINGS.URL);
    }
    
    var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    
    var hour = now.getHours();
    var mm = now.getMonth()+1;
    var yyyy = now.getYear();
    
    var days = ['January','February','March','April','May','June','July',
    			'August','September','October','November','December'];
    var monthName = days[now.getMonth()];
    
    var date = now.getDate();
    var daysInMonth = new Date(yyyy, mm, 0).getDate();
    
    var prevMonth = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    prevMonth.setMonth(prevMonth.getMonth() - 1)
    var prevMonthName = days[prevMonth.getMonth()];
    
    var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy').toUpperCase();
    
    var spreadsheet = SpreadsheetApp.openByUrl(SETTINGS.URL);
    var kpiSheet = spreadsheet.getSheetByName('KPI Report');
    var sheet = spreadsheet.getSheetByName('Daily Report');
    
    sheet.getRange(1,1,1,1).setValue(SETTINGS.LABEL + ' - AdWords Monthly Performance Report');
    sheet.getRange(2,1,1,1).setValue('Last Compiled On: ' + today);      
    kpiSheet.getRange(2,1,1,1).setValue('Last Compiled On: ' + today);  
    
    sheet.getRange('E4').setValue(prevMonthName);
    sheet.getRange('H4').setValue(prevMonthName);
    sheet.getRange('K4').setValue(prevMonthName);
    sheet.getRange('S4').setValue(prevMonthName);
    sheet.getRange('W4').setValue(prevMonthName);
    
    
    sheet.getRange('F4').setValue(monthName);
    sheet.getRange('I4').setValue(monthName);
    sheet.getRange('L4').setValue(monthName);
    sheet.getRange('T4').setValue(monthName);
    sheet.getRange('X4').setValue(monthName);
    
    kpiSheet.getRange('F4').setValue(monthName);
    kpiSheet.getRange('I4').setValue(monthName);
    kpiSheet.getRange('L4').setValue(monthName);
    kpiSheet.getRange('N4').setValue(monthName);
    kpiSheet.getRange('R4').setValue(monthName);
    
    var factor = parseFloat(date - 1 + hour/24);
    
    var total = sheet.getLastRow();
    var row = 0;
    
    for(var i = 6; i <= total; i++) {
      var cName = sheet.getRange("B"+i).getValue();
      if(cName == SETTINGS.CLIENT) {
        row = i;
        break;
      }
    }
    
    var totalKpi = kpiSheet.getLastRow();
    var rowKpi = 0;
    for(var i = 6; i <= totalKpi; i++) {
      var cName = kpiSheet.getRange("B"+i).getValue();
      if(cName == SETTINGS.CLIENT) {
        rowKpi = i;
        break;
      }
    }
    
    
    var statsNow = AdWordsApp.currentAccount().getStatsFor('THIS_MONTH');
    var clicksNow = statsNow.getClicks();
    var costNow = statsNow.getCost();
    var conversionsNow = statsNow.getConversions();
    var cpcNow = statsNow.getAverageCpc();
    
    var statsPrev = AdWordsApp.currentAccount().getStatsFor('LAST_MONTH');
    var clicksPrev = statsPrev.getClicks();
    var costPrev = statsPrev.getCost();
    var conversionsPrev = statsPrev.getConversions();
    var cpcPrev = statsPrev.getAverageCpc();
    
    var yestCost = AdWordsApp.currentAccount().getStatsFor('YESTERDAY').getCost();
    
    var clicksToLastMonth = clicksPrev == 0 ? 0 : (clicksNow/clicksPrev)/(factor/daysInMonth);
    var convToLastMonth = conversionsPrev == 0 ? 0 : (conversionsNow/conversionsPrev)/(factor/daysInMonth);
    var remainigBudgetAtSpends = SETTINGS.MONTHLY_BUDGET - ((daysInMonth*costNow)/factor);
    var dailyAvailableBudget = (SETTINGS.MONTHLY_BUDGET - costNow)/(daysInMonth - factor);
    var avgDailySpends = costNow/factor;
    var plusMinusDailyAvailable = dailyAvailableBudget - avgDailySpends;
    
    var cpaPrev = (conversionsPrev == 0) ? 'NA' : costPrev/conversionsPrev;
    var cpaNow = (conversionsNow == 0) ? 'NA' : costNow/conversionsNow;
    
    if(cpaNow == 'NA' || cpaPrev == 'NA') {
      var changeCpa = 'NA';
    } else {
      var changeCpa = (cpaNow - cpaPrev);
    }
    
    if(row == 0) {  
      row = total + 1; 
      var statsRow = [SETTINGS.MANAGER, SETTINGS.CLIENT, '', SETTINGS.MONTHLY_BUDGET, clicksPrev, clicksNow, clicksToLastMonth,
                      cpcPrev, cpcNow, (cpcNow - cpcPrev), costPrev, costNow, (SETTINGS.MONTHLY_BUDGET - costNow),
                      remainigBudgetAtSpends, dailyAvailableBudget, avgDailySpends, plusMinusDailyAvailable, 
                      yestCost, conversionsPrev, conversionsNow, (daysInMonth*conversionsNow)/factor, convToLastMonth,
                      cpaPrev, cpaNow, changeCpa, now, SETTINGS.CPA_TARGET, SETTINGS.DAILY_BUDGET];
      sheet.getRange(row,1,1,statsRow.length).setValues([statsRow]); 
    } else {
      var statsRow = [SETTINGS.MONTHLY_BUDGET, clicksPrev, clicksNow, clicksToLastMonth,
                      cpcPrev, cpcNow, (cpcNow - cpcPrev), costPrev, costNow, (SETTINGS.MONTHLY_BUDGET - costNow),
                      remainigBudgetAtSpends, dailyAvailableBudget, avgDailySpends, plusMinusDailyAvailable, 
                      yestCost, conversionsPrev, conversionsNow, (daysInMonth*conversionsNow)/factor, convToLastMonth,
                      cpaPrev, cpaNow, changeCpa, now, SETTINGS.CPA_TARGET, SETTINGS.DAILY_BUDGET];
      sheet.getRange(row,1,1,2).setValues([[SETTINGS.MANAGER, SETTINGS.CLIENT]]);
      sheet.getRange(row,4,1,statsRow.length).setValues([statsRow]);
    }
    
    if(rowKpi == 0) {
      rowKpi = kpiSheet.getLastRow() + 1;
    }
    
    kpiSheet.getRange(rowKpi,1,1,2).setValues([[SETTINGS.MANAGER, SETTINGS.CLIENT]]);
    kpiSheet.getRange(rowKpi,4,1,1).setValue(SETTINGS.MONTHLY_BUDGET);
    kpiSheet.getRange(rowKpi,6,1,1).setValue(clicksNow);
    kpiSheet.getRange(rowKpi,9,1,1).setValue(cpcNow);
    kpiSheet.getRange(rowKpi,12,1,1).setValue(costNow);
    kpiSheet.getRange(rowKpi,14,1,2).setValues([[conversionsNow,(daysInMonth*conversionsNow)/factor]]);
    kpiSheet.getRange(rowKpi,18,1,1).setValue(cpaNow);
    kpiSheet.getRange(rowKpi,20,1,9).setValues([[(SETTINGS.MONTHLY_BUDGET - costNow),remainigBudgetAtSpends, 
                                                 dailyAvailableBudget, avgDailySpends, plusMinusDailyAvailable, 
                                                 yestCost,now, SETTINGS.CPA_TARGET, SETTINGS.DAILY_BUDGET]]);
    
    kpiSheet.getRange(rowKpi,7,1,1).setFormula('=IFERROR((F'+rowKpi+'/E'+rowKpi+')/'+(factor/daysInMonth)+', "NA")');    
    kpiSheet.getRange(rowKpi,16,1,1).setFormula('=IFERROR((N'+rowKpi+'/M'+rowKpi+')/'+(factor/daysInMonth)+', "NA")');    
    kpiSheet.getRange(rowKpi,10,1,1).setFormula('=IFERROR(I'+rowKpi+'-H'+rowKpi+', "NA")'); 
    kpiSheet.getRange(rowKpi,19,1,1).setFormula('=IFERROR(R'+rowKpi+'-Q'+rowKpi+', "NA")'); 
    
    
    //sheet.getRange(row,27,1,2).setNumberFormat("#,##0");
    var rowData = sheet.getRange(row,1,1,26).getValues();
    var note = getNote(rowData[0]);
    sheet.getRange(row,2,1,1).setNote(note);
    rowData[0].push(AdWordsApp.currentAccount().getCustomerId());
    rowData[0].push(SETTINGS.CPA_TARGET);
    rowData[0].push(SETTINGS.DAILY_BUDGET);
    //masterS.getRange(rowNumMaster,1,1,rowData[0].length).setValues(rowData);
    
    if(SETTINGS.IS_RUNNING_IN_PARLLEL_MODE) { formatSheet(sheet) };
  }
}

function compileDailyReportNoKPI(SETTINGS) {
  var URLS = [SETTINGS.REPORT_URL];
  
  //if(SETTINGS.FLAG.toUpperCase() != 'Y'){ info('Script turned off for the account. Exiting.'); return; }
  
  
  var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy').toUpperCase();
  
  for(var l in URLS) {      
    SETTINGS.URL = URLS[l];      
    if(SETTINGS.URL.length < 5) {
      info('Spreadsheet Url Not found. Exiting');
      continue;
    } else {
      info(SETTINGS.URL);
    }
    
    var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    
    var hour = now.getHours();
    var mm = now.getMonth()+1;
    var yyyy = now.getYear();
    
    var days = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var monthName = days[now.getMonth()];
    
    var date = now.getDate();
    var daysInMonth = new Date(yyyy, mm, 0).getDate();
    
    var prevMonth = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    prevMonth.setMonth(prevMonth.getMonth() - 1)
    var prevMonthName = days[prevMonth.getMonth()];
    
    var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy').toUpperCase();
    
    var spreadsheet = SpreadsheetApp.openByUrl(SETTINGS.URL);
    var sheet = spreadsheet.getSheetByName('Daily Report');
    sheet.getRange(1,1,1,1).setValue(SETTINGS.LABEL + ' - AdWords Monthly Performance Report');
    sheet.getRange(2,1,1,1).setValue('Last Compiled On: ' + today);      
    
    sheet.getRange('E4').setValue(prevMonthName);
    sheet.getRange('H4').setValue(prevMonthName);
    sheet.getRange('K4').setValue(prevMonthName);
    sheet.getRange('S4').setValue(prevMonthName);
    sheet.getRange('W4').setValue(prevMonthName);
    
    sheet.getRange('F4').setValue(monthName);
    sheet.getRange('I4').setValue(monthName);
    sheet.getRange('L4').setValue(monthName);
    sheet.getRange('T4').setValue(monthName);
    sheet.getRange('X4').setValue(monthName);
    
    
    var factor = parseFloat(date - 1 + hour/24);
    
    var total = sheet.getLastRow();
    var row = 0;
    
    for(var i = 6; i <= total; i++) {
      var cName = sheet.getRange("B"+i).getValue();
      if(cName == SETTINGS.CLIENT) {
        row = i;
        break;
      }
    }
    
    
    var statsNow = AdWordsApp.currentAccount().getStatsFor('THIS_MONTH');
    var clicksNow = statsNow.getClicks();
    var costNow = statsNow.getCost();
    var conversionsNow = statsNow.getConversions();
    var cpcNow = statsNow.getAverageCpc();
    
    var statsPrev = AdWordsApp.currentAccount().getStatsFor('LAST_MONTH');
    var clicksPrev = statsPrev.getClicks();
    var costPrev = statsPrev.getCost();
    var conversionsPrev = statsPrev.getConversions();
    var cpcPrev = statsPrev.getAverageCpc();
    
    var yestCost = AdWordsApp.currentAccount().getStatsFor('YESTERDAY').getCost();
    
    var clicksToLastMonth = clicksPrev == 0 ? 0 : (clicksNow/clicksPrev)/(factor/daysInMonth);
    var convToLastMonth = conversionsPrev == 0 ? 0 : (conversionsNow/conversionsPrev)/(factor/daysInMonth);
    var remainigBudgetAtSpends = SETTINGS.MONTHLY_BUDGET - ((daysInMonth*costNow)/factor);
    var dailyAvailableBudget = (SETTINGS.MONTHLY_BUDGET - costNow)/(daysInMonth - factor);
    var avgDailySpends = costNow/factor;
    var plusMinusDailyAvailable = dailyAvailableBudget - avgDailySpends;
    
    var cpaPrev = (conversionsPrev == 0) ? 'NA' : costPrev/conversionsPrev;
    var cpaNow = (conversionsNow == 0) ? 'NA' : costNow/conversionsNow;
    
    if(cpaNow == 'NA' || cpaPrev == 'NA') {
      var changeCpa = 'NA';
    } else {
      var changeCpa = (cpaNow - cpaPrev);
    }
    
    if(row == 0) {  
      row = total + 1; 
      var statsRow = [SETTINGS.MANAGER, SETTINGS.CLIENT, '', SETTINGS.MONTHLY_BUDGET, clicksPrev, clicksNow, clicksToLastMonth,
                      cpcPrev, cpcNow, (cpcNow - cpcPrev), costPrev, costNow, (SETTINGS.MONTHLY_BUDGET - costNow),
                      remainigBudgetAtSpends, dailyAvailableBudget, avgDailySpends, plusMinusDailyAvailable, 
                      yestCost, conversionsPrev, conversionsNow, (daysInMonth*conversionsNow)/factor, convToLastMonth,
                      cpaPrev, cpaNow, changeCpa, now, SETTINGS.CPA_TARGET, SETTINGS.DAILY_BUDGET];
      sheet.getRange(row,1,1,statsRow.length).setValues([statsRow]); 
    } else {
      var statsRow = [SETTINGS.MONTHLY_BUDGET, clicksPrev, clicksNow, clicksToLastMonth,
                      cpcPrev, cpcNow, (cpcNow - cpcPrev), costPrev, costNow, (SETTINGS.MONTHLY_BUDGET - costNow),
                      remainigBudgetAtSpends, dailyAvailableBudget, avgDailySpends, plusMinusDailyAvailable, 
                      yestCost, conversionsPrev, conversionsNow, (daysInMonth*conversionsNow)/factor, convToLastMonth,
                      cpaPrev, cpaNow, changeCpa, now, SETTINGS.CPA_TARGET, SETTINGS.DAILY_BUDGET];
      sheet.getRange(row,1,1,2).setValues([[SETTINGS.MANAGER, SETTINGS.CLIENT]]);
      sheet.getRange(row,4,1,statsRow.length).setValues([statsRow]);
    }
    
    //sheet.getRange(row,27,1,2).setNumberFormat("#,##0");
    var rowData = sheet.getRange(row,1,1,26).getValues();
    var note = getNote(rowData[0]);
    sheet.getRange(row,2,1,1).setNote(note);
    rowData[0].push(AdWordsApp.currentAccount().getCustomerId());
    rowData[0].push(SETTINGS.CPA_TARGET);
    rowData[0].push(SETTINGS.DAILY_BUDGET);
    //masterS.getRange(rowNumMaster,1,1,rowData[0].length).setValues(rowData);
    
    formatSheet(sheet);
    compileSummary(sheet,spreadsheet);
    compileConversionReport(SETTINGS.URL);
    compileEcommerceReport(SETTINGS.URL);
  }
}

function getNote(data) {    
  var note = '';
  
  var budgetVariance = 'underspend';
  if(data[13] < 0) { budgetVariance = 'overspend'; }
  
  var cpaChange = ' less';
  if(data[24] > 0) { cpaChange = ' more'; }
  
  var cpcChange = ' less';
  if(data[9] > 0) { cpcChange = ' more'; }
  
  try {
    note = 'Conversions this month are expected to hit ' + (100*data[21]).toFixed(2) + '% of last month conversions, and the CPA is ' + (Math.abs(data[24])).toFixed(2) + cpaChange +', \
visitor numbers are ' + (100*data[6]).toFixed(2) + '% at the same time last month. \
\n\nCPC is '+(data[9]).toFixed(2)+'p' +cpcChange+' than last month and you will ' + budgetVariance+ ' by ' + (data[13]).toFixed(2) + '.'; 
    
  } catch(ex) {
    info(ex);
  }
  
  return note;    
}

function formatSheet(sheet){
  
  var numRows = sheet.getLastRow();
  if(numRows < 1) { return; }
  
  sheet.getRange(6,1,numRows,3).setBackground('#d9ead3');
  sheet.getRange(6,4,numRows,1).setBackground('#d9ead3').setNumberFormat("#,##0.00"); // Budget
  sheet.getRange(6,5,numRows,1).setBackground('#ffffcc').setNumberFormat("#,##0");
  sheet.getRange(6,6,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0");
  sheet.getRange(6,7,numRows,1).setBackground('#efefef').setNumberFormat("0.00%"); //Clicks @ Current Rate
  sheet.getRange(6,8,numRows,1).setBackground('#ffffcc').setNumberFormat("#,##0.00"); //LM CPC
  sheet.getRange(6,9,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0.00"); // TM CPC
  sheet.getRange(6,10,numRows,1).setBackground('#efefef').setNumberFormat("#,##0.00"); //CPC DIFF
  sheet.getRange(6,11,numRows,1).setBackground('#ffffcc').setNumberFormat("#,##0.00"); //LM Spends
  sheet.getRange(6,12,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0.00"); //TM Spends
  sheet.getRange(6,13,numRows,5).setBackground('#efefef').setNumberFormat("#,##0.00");
  sheet.getRange(6,18,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0.00"); // Yest Spends
  sheet.getRange(6,19,numRows,1).setBackground('#ffffcc').setNumberFormat("#,##0");  // LM Conv
  sheet.getRange(6,20,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0");  // TM Conv
  sheet.getRange(6,21,numRows,1).setBackground('#efefef').setNumberFormat("#,##0");  // Conv @ Current Rate
  sheet.getRange(6,22,numRows,1).setBackground('#efefef').setNumberFormat("0.00%"); //Conv % of Last Month
  sheet.getRange(6,23,numRows,1).setBackground('#ffffcc').setNumberFormat("#,##0.00"); //LM CPA
  sheet.getRange(6,24,numRows,1).setBackground('#ffe599').setNumberFormat("#,##0.00"); //TM CPA
  sheet.getRange(6,25,numRows,1).setBackground('#efefef').setNumberFormat("#,##0.00"); //CPA DIFF
  sheet.getRange(6,26,numRows,1).setBackground('#d9ead3');
  
  var sheetData = sheet.getDataRange().getValues();
  
  for(var k in sheetData) {
    if(k < 5) { continue; }
    if(sheetData[k][1] == '') { continue; }
    var row = parseInt(k,10) + 1;
    
    var col = 10; //Change in CPC
    var cpcChange = sheetData[k][col-1];
    if(cpcChange > 0){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }
    
    var col = 25; // Change in CPA
    var cpaChange = sheetData[k][col-1];
    if(cpaChange > 0){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b'); //Red
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');  //Grey
    }
    
    var col = 7; // Clicks % 
    var clickPercent = sheetData[k][col-1];
    if(clickPercent < 0.8){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else if(clickPercent > 1.2){
      sheet.getRange(row,col,1,1).setBackground('#b6d7a8');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }
    
    var col = 13;  // Remaining Budget
    var remainingBudget = sheetData[k][col-1];
    if(remainingBudget < 0){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }
    
    var col = 14;  // Remianing Budget PPC
    var remainingBudgetPPC = sheetData[k][col-1];
    if(remainingBudgetPPC < 0){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }	  
    
    var col = 17; // Available Budget 
    var availableDailyBudget = sheetData[k][col-1];
    if(availableDailyBudget < 0){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }
    
    var col = 22; // Conv %
    var convPercentOnLastMonth = sheetData[k][col-1];
    if(convPercentOnLastMonth > 1.20){
      sheet.getRange(row,col,1,1).setBackground('#b6d7a8');
    } else if(convPercentOnLastMonth < 0.80){
      sheet.getRange(row,col,1,1).setBackground('#dd7e6b');
    } else {
      sheet.getRange(row,col,1,1).setBackground('#d9d9d9');
    }
  }	
}


function compileSummary(reportSheet, spreadsheet) {  
  info('Compiling Summary Report');
  var data = reportSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  
  var visitorsPct = [];
  var cpcChange = [];
  var cpaChange = [];
  var remainingSpends = [];
  var conversionsPct = [];
  
  for(var k in data) {
    if(data[k][6] != 'NA' && data[k][6] != 0 && data[k][6]!='#NUM!') { visitorsPct.push([data[k][1],data[k][6]]); }
    if(data[k][9] != 'NA' && data[k][9] != 0 && data[k][9]!='#NUM!') { cpcChange.push([data[k][1],data[k][9]]); }
    if(data[k][13] != 'NA' && data[k][13] != 0 && data[k][13]!='#NUM!') { remainingSpends.push([data[k][1],data[k][13]]); }
    if(data[k][21] != 'NA' && data[k][21] != 0 && data[k][21]!='#NUM!') { conversionsPct.push([data[k][1],data[k][21]]); }
    if(data[k][24] != 'NA' && data[k][24] != 0 && data[k][24]!='#NUM!') { cpaChange.push([data[k][1],data[k][24]]); }    
  }
  
  visitorsPct.sort(function(a,b) {
    return b[1] - a[1];
  });
  
  cpcChange.sort(function(a,b) {
    return b[1] - a[1];
  });
  
  remainingSpends.sort(function(a,b) {
    return b[1] - a[1];
  });
  
  conversionsPct.sort(function(a,b) {
    return b[1] - a[1];
  });
  
  cpaChange.sort(function(a,b) {
    return b[1] - a[1];
  });
  
  
  var topVisitors = [['Top 5 with Biggest Increase in visitors', ''], ['Client', '% Visitors to Last Month']]; 
  for(var k = 0; k < 5; k++) {   
    if(!visitorsPct[k]) {
      topVisitors.push(['','']);    
      continue;
    }
    topVisitors.push(visitorsPct[k]);    
  }
  
  var topCpc = [['Biggest Increases in CPC', ''], ['Client','CPC Diff']];
  for(var k = 0; k < 5; k++) {   
    if(!cpcChange[k]) {
      topCpc.push(['','']);    
      continue;
    }
    topCpc.push(cpcChange[k]);    
  }  
  
  var topSpends = [['Biggest Underspends', ''], ['Client','Remaining Spends']];
  for(var k = 0; k < 5; k++) { 
    if(!remainingSpends[k]) {
      topSpends.push(['','']);    
      continue;
    }
    topSpends.push(remainingSpends[k]);    
  }
  
  var topConv = [['Biggest Gain in Expected Conversions', ''], ['Client','% Conversions to Last Month']]; 
  for(var k = 0; k < 5; k++) {    
    if(!conversionsPct[k]) {
      topConv.push(['','']);    
      continue;
    }
    topConv.push(conversionsPct[k]);    
  }
  
  var topCpa = [['Biggest Change in CPA',''],['Client','CPC Change']];
  for(var k = 0; k < 5; k++) {    
    if(!cpaChange[k]) {
      topCpa.push(['','']);    
      continue;
    }
    topCpa.push(cpaChange[k]);    
  }
  
  var bottomVisitors = [['Worst 5 with biggest decrease in visitors', ''], ['Client','% Visitors to Last Month']]; 
  var size = visitorsPct.length-1;
  for(var k = size; k > size - 5; k--) {
    if(!visitorsPct[k]) {
      bottomVisitors.push(['','']);    
      continue;
    }
    bottomVisitors.push(visitorsPct[k]);    
  }
  
  var bottomCpc = [['Best Decreases in CPC',''],['Client','CPC Diff']];
  size = cpcChange.length-1;
  for(var k = size; k > size - 5; k--) { 
    if(!cpcChange[k]) {
      bottomCpc.push(['','']);    
      continue;
    }
    bottomCpc.push(cpcChange[k]);    
  }  
  
  var bottomSpends = [['Biggest Overspends', ''], ['Client','Remaining Spends']];
  size = remainingSpends.length-1;
  for(var k = size; k > size - 5; k--) {  
    if(!remainingSpends[k]) {
      bottomSpends.push(['','']);    
      continue;
    }
    bottomSpends.push(remainingSpends[k]);    
  }
  
  var bottomConv = [['Worst drop in expected conversions', ''], ['Client','% Conversions to Last Month']]; 
  size = conversionsPct.length-1;
  for(var k = size; k > size - 5; k--) {
    if(!conversionsPct[k]) {
      bottomConv.push(['','']);    
      continue;
    }
    bottomConv.push(conversionsPct[k]);    
  }
  
  var bottomCpa = [['Best Change in CPA',''],['Client','CPA Change']];
  size = cpaChange.length-1;
  for(var k = size; k > size - 5; k--) {
    if(!cpaChange[k]) {
      bottomCpa.push(['','']);    
      continue;
    }
    bottomCpa.push(cpaChange[k]);    
  }    
  
  var sheet = getSheet('Summary Report', spreadsheet);
  sheet.getRange(1,1,topVisitors.length,2).setValues(topVisitors).setNumberFormat("0.00%");    
  sheet.getRange(1,4,topCpc.length,2).setValues(topCpc).setNumberFormat("#,##0.00");
  sheet.getRange(1,7,topSpends.length,2).setValues(topSpends).setNumberFormat("#,##0.00");
  sheet.getRange(9,1,bottomVisitors.length,2).setValues(bottomVisitors).setNumberFormat("0.00%");
  sheet.getRange(9,4,bottomCpc.length,2).setValues(bottomCpc).setNumberFormat("#,##0.00");
  sheet.getRange(9,7,bottomSpends.length,2).setValues(bottomSpends).setNumberFormat("#,##0.00");
  sheet.getRange(17,1,topConv.length,2).setValues(topConv).setNumberFormat("0.00%");
  sheet.getRange(17,4,topCpa.length,2).setValues(topCpa).setNumberFormat("#,##0.00");
  sheet.getRange(25,1,bottomConv.length,2).setValues(bottomConv).setNumberFormat("0.00%");
  sheet.getRange(25,4,bottomCpa.length,2).setValues(bottomCpa).setNumberFormat("#,##0.00");
  
  sheet.getRange(1,1,1,2).merge().setBackground('#6aa84f').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(1,4,1,2).merge().setBackground('#6aa84f').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(1,7,1,2).merge().setBackground('#6aa84f').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(9,1,1,2).merge().setBackground('#e6b8af').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(9,4,1,2).merge().setBackground('#e6b8af').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(9,7,1,2).merge().setBackground('#e6b8af').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(17,1,1,2).merge().setBackground('#6aa84f').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(17,4,1,2).merge().setBackground('#e6b8af').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(25,1,1,2).merge().setBackground('#e6b8af').setFontWeight('bold').setHorizontalAlignment("center");
  sheet.getRange(25,4,1,2).merge().setBackground('#6aa84f').setFontWeight('bold').setHorizontalAlignment("center");
  
  sheet.getRange(2,1,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(2,4,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(2,7,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(10,1,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(10,4,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(10,7,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(18,1,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(18,4,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(26,1,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');
  sheet.getRange(26,4,1,2).setBackgrounds([['#cfe2f3','#cfe2f3']]).setFontWeight('bold');   
  
  sheet.getDataRange().setFontFamily('Calibri').setFontSize(10);
  deleteExtraRowsCols(sheet);   
}

function getSheet(name,ss) {
  var sheet = ss.getSheetByName(name);
  if(!sheet) { sheet = ss.insertSheet(name); }
  
  return sheet;
}

function deleteExtraRowsCols(sheet) {
  if((sheet.getMaxColumns() - sheet.getLastColumn()) > 0) {
    sheet.deleteColumns(sheet.getLastColumn()+1, sheet.getMaxColumns() - sheet.getLastColumn());
  }
  
  if((sheet.getMaxRows() - sheet.getLastRow()) > 0) {
    sheet.deleteRows(sheet.getLastRow()+1, sheet.getMaxRows() - sheet.getLastRow());
  }
}

function compileConversionReport(url) {
  info('Compiling Conversion Report');
  var sheet = getSheetByName(url, 'Conversion Report');
  
  if(sheet.getRange(4,2,1,1).getValue() != 'Today') {
    sheet.clear();
    setupConversionReportSheet(sheet);
  }
  
  var yesterday = getAdWordsFormattedDate(1, 'yyyyMMdd');
  var twoDaysBack = getAdWordsFormattedDate(2, 'yyyyMMdd');
  var threeDaysBack = getAdWordsFormattedDate(3, 'yyyyMMdd');
  
  var todayStats = AdWordsApp.currentAccount().getStatsFor('TODAY');
  var yesterdayStats = AdWordsApp.currentAccount().getStatsFor(yesterday,yesterday);
  var twoDaysStats = AdWordsApp.currentAccount().getStatsFor(twoDaysBack,twoDaysBack);
  var threeDaysStats = AdWordsApp.currentAccount().getStatsFor(threeDaysBack,threeDaysBack);
  var thirtyDaysStats = AdWordsApp.currentAccount().getStatsFor('LAST_30_DAYS');
  
  var todayConv = todayStats.getConversions();
  var yesterdayConv = yesterdayStats.getConversions();
  var twoDayConv = twoDaysStats.getConversions();
  var threeDayConv = threeDaysStats.getConversions();
  var thirtyDayConv = Math.round(thirtyDaysStats.getConversions()/30);
  
  var data = [[todayConv, yesterdayConv, twoDayConv, threeDayConv, thirtyDayConv,
               (todayConv == 0 ? 0 : (todayStats.getCost() / todayConv)),
               (yesterdayConv == 0 ? 0 : (yesterdayStats.getCost() / yesterdayConv)),
               (twoDayConv == 0 ? 0 : (twoDaysStats.getCost() / twoDayConv)),
               (threeDayConv == 0 ? 0 : (threeDaysStats.getCost() / threeDayConv)),
               (thirtyDayConv == 0 ? 0 : (thirtyDaysStats.getCost() / thirtyDayConv))					 
              ]];  
  
  
  var reportRow = getReportRow(sheet,AdWordsApp.currentAccount().getName());
  sheet.getRange(reportRow, 2, data.length, data[0].length).setValues(data);
  sheet.getRange(reportRow, 2, data.length, 5).setNumberFormat("#,##0"); // Conversions
  sheet.getRange(reportRow, 7, data.length, 5).setNumberFormat("#,##0.00"); // CPA    
  
  var dateHead = [[getAdWordsFormattedDate(2, 'MMM dd, yyyy'), getAdWordsFormattedDate(3, 'MMM dd, yyyy')]]
  sheet.getRange(4,4,1,2).setValues(dateHead);
  sheet.getRange(4,9,1,2).setValues(dateHead);
  
  sheet.getDataRange().setFontFamily('Calibri');
  
  if(sheet.getLastColumn() > 11) {
    var numCols = sheet.getLastColumn() - 11;
    sheet.getRange(1,12,sheet.getLastRow(),numCols).clear();
  }
}

function compileEcommerceReport(url) {
  info('Compiling Ecommerce Report');
  var sheet = getSheetByName(url, 'Ecommerce Report');
  
  var analyticsData = getDataForAnalytics();
  var data = [analyticsData.row]
  
  if(sheet.getRange(4,2,1,1).getValue() != 'Today') {
    setupEcommerceReportSheet(sheet);
  }
  
  sheet.getRange(4,2,1,analyticsData.headerRow.length).setValues([analyticsData.headerRow]).setFontWeight('bold').setBackground('#ffe599').setNumberFormat('@STRING@').setBorder(true,true,true,true,true,true);
  
  var reportRow = getReportRow(sheet,AdWordsApp.currentAccount().getName());
  sheet.getRange(reportRow, 2, data.length, data[0].length).setValues(data);  
  sheet.getRange(reportRow,2,data.length,9).setNumberFormat("#,##0.00"); // Revenue
  sheet.getRange(reportRow,11,data.length,9).setNumberFormat("#,##0"); // Trasnsactions  
  
  if(analyticsData.backgrounds.length > 0) {
    sheet.getRange(reportRow, 3, analyticsData.backgrounds.length, analyticsData.backgrounds[0].length).setBackgrounds(analyticsData.backgrounds);  
  }
  
  sheet.getDataRange().setFontFamily('Calibri');
}

function getDataForAnalytics() {
  var dataSheet = SpreadsheetApp.openByUrl(ANALYTICS_STATS_REPORT).getSheetByName(ANALYTICS_TAB_NAME);
  var data = dataSheet.getDataRange().getValues();
  data.shift();
  var header = data.shift();
  header.shift();
  header.shift();
  
  var accName = AdWordsApp.currentAccount().getName();
  
  var rNum = -1;
  var row = [];
  for(var k in data) {
    if(data[k][0] == accName) {
      data[k].shift();
      data[k].shift();
      row = data[k];
      rNum = parseInt(k,10) + 3;
      break;
    }        
  }
  
  var backgrounds = [];
  if(rNum != -1) {
    backgrounds = dataSheet.getRange(rNum, 4, 1, 3).getBackgrounds();
  }
  
  var max = 23;
  if(row.length == 0) {
    for(var j=0; j<max; j++) {
      row.push('');
    }       
  }
  
  return { row: row, headerRow: header, backgrounds: backgrounds }
}

function getReportRow(sheet,accName) {
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  
  var reportRow = -1;
  for(var k in data) {
    if(data[k][0] != accName) { continue; }
    reportRow = parseInt(k,10)+5;
  }
  
  if(reportRow != -1) { return reportRow; }
  reportRow = sheet.getLastRow()+1;
  sheet.getRange(reportRow,1,1,1).setValue(accName);
  return reportRow;
}

function getSheetByName(url, name) {
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = ss.getSheetByName(name);
  if(sheet) { return sheet; }
  
  sheet = ss.insertSheet(name);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
  return sheet;
}

function setupConversionReportSheet(sheet) {
  
  var header = ['Client name','Today','Yesterday',getAdWordsFormattedDate(2, 'MMM dd, yyyy'), getAdWordsFormattedDate(3, 'MMM dd, yyyy'),
                '30 Day Avg','Today','Yesterday',getAdWordsFormattedDate(2, 'MMM dd, yyyy'), getAdWordsFormattedDate(3, 'MMM dd, yyyy'),'30 Day Avg'];
  sheet.getRange(4,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#ffe599').setNumberFormat('@STRING@').setBorder(true,true,true,true,true,true);
  sheet.getRange(3,2,1,1).setValue('Converted Clicks');
  sheet.getRange(3,2,1,5).merge().setBackground('#d0e0e3');
  sheet.getRange(3,7,1,1).setValue('Cost Per Converted Click');
  sheet.getRange(3,7,1,5).merge().setBackground('#d9ead3');
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
  
  if((sheet.getMaxColumns() - sheet.getLastColumn()) > 0) {
    sheet.deleteColumns(sheet.getLastColumn()+1, sheet.getMaxColumns() - sheet.getLastColumn());
  }
}

function setupEcommerceReportSheet(sheet) {
  
  if(sheet.getLastColumn() > 1) {
    sheet.getRange(3,1,1,sheet.getLastColumn()).breakApart();
  }
  
  sheet.getRange(3,2,1,1).setValue('e-commerce Revenue');
  sheet.getRange(3,2,1,9).merge().setBackground('#f4cccc');
  sheet.getRange(3,11,1,1).setValue('e-commerce number of Transactions');
  sheet.getRange(3,11,1,9).merge().setBackground('#d9d2e9');
  sheet.getRange(3,20,1,1).setValue('Goal Completions');
  sheet.getRange(3,20,1,3).merge().setBackground('#d9ead3');
  
  if((sheet.getMaxColumns() - sheet.getLastColumn()) > 0) {
    sheet.deleteColumns(sheet.getLastColumn()+1, sheet.getMaxColumns() - sheet.getLastColumn());
  }      
}

function compileResults() {
  Logger.log('Finished');
}

function readAccountSettings() {
  var masterSS = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL);
  var INPUT_MAP = {};
  
  var data = masterSS.getSheetByName(DASHBOARD_URLS_TAB).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][1]) { continue; }
    var ss = SpreadsheetApp.openByUrl(data[k][1]);
    var sheet = ss.getSheetByName('Account Inputs');
    if(!sheet) { continue; }
    
    var inputData = sheet.getDataRange().getValues();
    var scriptNameHeader = inputData.shift();
    var inputHeader = inputData.shift();
    
    for(var j in inputData){
      var SETTINGS = new Object();
      for(var l in inputHeader) {
        SETTINGS[inputHeader[l]] = inputData[j][l];
      }
      INPUT_MAP[inputData[j][0]] = SETTINGS;
    }
  }
  
  return INPUT_MAP;
}


function info(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);
}

/**
* Get AdWords Formatted date for n days back
* @param {int} d - Numer of days to go back for start/end date
* @return {String} - Formatted date yyyyMMdd
**/
function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}




function compileWeeklyReport(SETTINGS) {
  
  var URLS = [SETTINGS.REPORT_URL];
  
  //if(SETTINGS.FLAG.toUpperCase() != 'Y'){ info('Script turned off for the account. Exiting.'); return; }
  
  var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy').toUpperCase();
  var masterSS = SpreadsheetApp.openByUrl(MASTER_MANAGEMENT_URL);     
  var masterS = masterSS.getSheetByName('Weekly & Monthly Report');
  
  var rowNumMaster = getReportRowWeekly(masterS, AdWordsApp.currentAccount().getName());
  
  var weeklyStats = getStatsByWeek();
  var monthlyStats = getStatsByMonth();
  
  weeklyStats.clicksChange = weeklyStats.statsOld.getClicks() == 0 ? '-' : (100*((weeklyStats.statsNew.getClicks() - weeklyStats.statsOld.getClicks()) / weeklyStats.statsOld.getClicks()))+'%';
  weeklyStats.impressionsChange = weeklyStats.statsOld.getImpressions() == 0 ? '-' : (100*((weeklyStats.statsNew.getImpressions() - weeklyStats.statsOld.getImpressions()) / weeklyStats.statsOld.getImpressions()))+'%';
  weeklyStats.costChange = weeklyStats.statsOld.getCost() == 0 ? '-' : (100*((weeklyStats.statsNew.getCost() - weeklyStats.statsOld.getCost()) / weeklyStats.statsOld.getCost()))+'%';
  weeklyStats.cpcChange = weeklyStats.statsOld.getAverageCpc() == 0 ? '-' : (100*((weeklyStats.statsNew.getAverageCpc() - weeklyStats.statsOld.getAverageCpc()) / weeklyStats.statsOld.getAverageCpc()))+'%';
  weeklyStats.conversionsChange = weeklyStats.statsOld.getConversions() == 0 ? '-' : (100*((weeklyStats.statsNew.getConversions() - weeklyStats.statsOld.getConversions()) / weeklyStats.statsOld.getConversions()))+'%';
  
  weeklyStats.statsNew.cpa = weeklyStats.statsNew.getConversions() == 0 ? 0 : (weeklyStats.statsNew.getCost() / weeklyStats.statsNew.getConversions());
  weeklyStats.statsOld.cpa = weeklyStats.statsOld.getConversions() == 0 ? 0 : (weeklyStats.statsOld.getCost() / weeklyStats.statsOld.getConversions());
  
  weeklyStats.cpaChange = weeklyStats.statsOld.cpa == 0 ? '-' : (100*((weeklyStats.statsNew.cpa - weeklyStats.statsOld.cpa) / weeklyStats.statsOld.cpa))+'%';
  
  monthlyStats.clicksChange = monthlyStats.statsOld.getClicks() == 0 ? '-' : (100*((monthlyStats.statsNew.getClicks() - monthlyStats.statsOld.getClicks()) / monthlyStats.statsOld.getClicks()))+'%';
  monthlyStats.impressionsChange = monthlyStats.statsOld.getImpressions() == 0 ? '-' : (100*((monthlyStats.statsNew.getImpressions() - monthlyStats.statsOld.getImpressions()) / monthlyStats.statsOld.getImpressions()))+'%';
  monthlyStats.costChange = monthlyStats.statsOld.getCost() == 0 ? '-' : (100*((monthlyStats.statsNew.getCost() - monthlyStats.statsOld.getCost()) / monthlyStats.statsOld.getCost()))+'%';
  monthlyStats.cpcChange = monthlyStats.statsOld.getAverageCpc() == 0 ? '-' : (100*((monthlyStats.statsNew.getAverageCpc() - monthlyStats.statsOld.getAverageCpc()) / monthlyStats.statsOld.getAverageCpc()))+'%';
  monthlyStats.conversionsChange = monthlyStats.statsOld.getConversions() == 0 ? '-' : (100*((monthlyStats.statsNew.getConversions() - monthlyStats.statsOld.getConversions()) / monthlyStats.statsOld.getConversions()))+'%';
  
  monthlyStats.statsNew.cpa = monthlyStats.statsNew.getConversions() == 0 ? 0 : (monthlyStats.statsNew.getCost() / monthlyStats.statsNew.getConversions());
  monthlyStats.statsOld.cpa = monthlyStats.statsOld.getConversions() == 0 ? 0 : (monthlyStats.statsOld.getCost() / monthlyStats.statsOld.getConversions());
  
  monthlyStats.cpaChange = monthlyStats.statsOld.cpa == 0 ? '-' : (100*((monthlyStats.statsNew.cpa - monthlyStats.statsOld.cpa) / monthlyStats.statsOld.cpa))+'%';
  
  var row = [AdWordsApp.currentAccount().getName(), weeklyStats.statsNew.getClicks(), weeklyStats.clicksChange, 
             weeklyStats.statsNew.getImpressions(), weeklyStats.impressionsChange,     
             weeklyStats.statsNew.getCost(), weeklyStats.costChange,
             weeklyStats.statsNew.getAverageCpc(), weeklyStats.cpcChange,
             weeklyStats.statsNew.getConversions(), weeklyStats.conversionsChange,
             weeklyStats.statsNew.cpa, weeklyStats.cpaChange,
             monthlyStats.statsNew.getClicks(), monthlyStats.clicksChange, 
             monthlyStats.statsNew.getImpressions(), monthlyStats.impressionsChange, 	
             monthlyStats.statsNew.getCost(), monthlyStats.costChange,
             monthlyStats.statsNew.getAverageCpc(), monthlyStats.cpcChange,
             monthlyStats.statsNew.getConversions(), monthlyStats.conversionsChange,
             monthlyStats.statsNew.cpa, monthlyStats.cpaChange
            ];
  
  masterS.getRange(rowNumMaster, 1, 1, row.length).setValues([row]);
  
  masterS.getRange(1,2,1,1).setValue('Weekly: ' + weeklyStats.dateRange);      
  masterS.getRange(1,14,1,1).setValue('Monthly: ' + monthlyStats.dateRange); 
  masterS.getDataRange().setFontFamily('Calibri');
  
  var MSG = '';
  //if(LABEL == 'Isuru') {
  MSG = compileWeeklyReportEmail(weeklyStats,monthlyStats);
  //}
  
  for(var l in URLS) {      
    SETTINGS.URL = URLS[l];      
    if(SETTINGS.URL.length < 5) {
      info('Spreadsheet Url Not found. Exiting');
      continue;
    } else {
      info(SETTINGS.URL);
    }
    
    var spreadsheet = SpreadsheetApp.openByUrl(SETTINGS.URL);
    var sheet = getSheetByNameKey(spreadsheet, 'Weekly & Monthly Report', 1);
    
    sheet.getRange(1,2,1,1).setValue('Weekly: ' + weeklyStats.dateRange);      
    sheet.getRange(1,14,1,1).setValue('Monthly: ' + monthlyStats.dateRange); 
    
    var num = getReportRowWeekly(sheet, AdWordsApp.currentAccount().getName());
    
    sheet.getRange(num, 1, 1, row.length).setValues([row]);
    
    if(MSG) {
      var tempSheet = getSheetByNameKey(spreadsheet, 'Email Templates', 0);
      tempSheet.getRange(1,1,1,2).setValues([['Client Name', 'Email Template']]);
      tempSheet.setFrozenRows(1);
      var existingData = tempSheet.getDataRange().getValues();
      existingData.shift();
      
      var rNum = 0;
      for(var k in existingData) {
        if(existingData[k][0] == AdWordsApp.currentAccount().getName()) {
          rNum = parseInt(k,10)+2;
          break;
        }
      }
      
      if(!rNum) {
        rNum = tempSheet.getLastRow() + 1;
        tempSheet.getRange(rNum,1,1,1).setValue(AdWordsApp.currentAccount().getName());
      }
      
      tempSheet.getRange(rNum,2,1,1).setValue(MSG);
    }
    
  }
  
  sheet.getDataRange().setFontFamily('Calibri');
}

function compileWeeklyReportEmail(weeklyStats,monthlyStats) {
  
  var MSG = 'Hi,\n\nA weekly performance update for ' + AdWordsApp.currentAccount().getName() + ' from the ' + weeklyStats.startDateFormatted + ' - ' + weeklyStats.endDateFormatted +' compared to the week before.\n\nIn total there were ' + weeklyStats.statsNew.getClicks() + ' clicks with an average cost per click of £' + weeklyStats.statsNew.getAverageCpc() +'. ';
  
  var weeklyCpcChange = weeklyStats.statsOld.getAverageCpc() == 0 ? 0 : 100*((weeklyStats.statsNew.getAverageCpc() - weeklyStats.statsOld.getAverageCpc()) / weeklyStats.statsOld.getAverageCpc());
  
  if(weeklyCpcChange < 0) {    
    MSG += 'There was a ' + weeklyCpcChange.toFixed(2) + '% decrease in CPC. ';
  }
  
  MSG += 'This led to a total weekly spend of £' + weeklyStats.statsNew.getCost() + '.';
  
  var weeklyImpressionsChange = weeklyStats.statsOld.getImpressions() == 0 ? 0 : 100*((weeklyStats.statsNew.getImpressions() - weeklyStats.statsOld.getImpressions()) / weeklyStats.statsOld.getImpressions());
  
  if(weeklyImpressionsChange > 0) {
    MSG += '\nThere was more traffic to the account this past week with ' + weeklyImpressionsChange.toFixed(2) + '% more impressions.'
  }
  
  var weeklyConversionsChange = weeklyStats.statsOld.getConversions() == 0 ? 0 : 100*((weeklyStats.statsNew.getConversions() - weeklyStats.statsOld.getConversions()) / weeklyStats.statsOld.getConversions());
  
  if(weeklyConversionsChange > 0) {
    MSG += '\n\nThere was ' + weeklyStats.statsNew.getConversions() + ' conversions which was a ' + weeklyConversionsChange.toFixed(2) + '% increase from the week before. ';
    var weeklyCpaChange = weeklyStats.statsOld.cpa == 0 ? 0 : 100*((weeklyStats.statsNew.cpa - weeklyStats.statsOld.cpa) / weeklyStats.statsOld.cpa);
    
    if(weeklyCpaChange < 0) {
      MSG += 'This led to a ' + weeklyCpaChange.toFixed(2) + '% drop in the CPA which was averaging at £' + weeklyStats.statsNew.cpa.toFixed(2) + ' for the past week.';
    }
  }
  
  MSG += '\n\nSo far this month there has been ' + monthlyStats.statsNew.getClicks() + ' clicks with a total spend of £' + monthlyStats.statsNew.getCost() + '. In total there has been ' + monthlyStats.statsNew.getConversions() + ' conversions with a average cost per lead of £' + monthlyStats.statsNew.cpa.toFixed(2) + '.\n\nKind Regards'
  
  var sheet = SpreadsheetApp.openByUrl(EMAIL_TEMPLATE_URL).getSheetByName('Email Templates');
  var data = sheet.getDataRange().getValues();
  data.shift();
  
  var rNum = 0;
  for(var k in data) {
    if(data[k][0] == AdWordsApp.currentAccount().getName()) {
      rNum = parseInt(k,10)+2;
      break;
    }
  }
  
  if(!rNum) {
    rNum = sheet.getLastRow() + 1;
    sheet.getRange(rNum,1,1,1).setValue(AdWordsApp.currentAccount().getName());
  }
  
  sheet.getRange(rNum,2,1,1).setValue(MSG);
  
  return MSG;
}

function getStatsByMonth() {
  var STATS = {};
  
  var dummyDate = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  dummyDate.setDate(dummyDate.getDate()-1);
  dummyDate.setHours(12);
  dummyDate.setDate(0);
  var daysInMonth = dummyDate.getDate();
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  date.setDate(date.getDate()-1);
  date.setHours(12);
  
  STATS.endDate = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.endDateFormatted = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  while(date.getDate() > daysInMonth) {
    date.setDate(date.getDate()-1);
  }
  
  date.setMonth(date.getMonth()-1);
  STATS.endDateOld = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.endDateFormattedOld = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  
  date.setMonth(date.getMonth()+1);
  date.setDate(1);
  STATS.startDate = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.startDateFormatted = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  
  date.setMonth(date.getMonth()-1);
  STATS.startDateOld = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.startDateFormattedOld = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  STATS.dateRange = '(' + STATS.startDateFormatted + ' - ' + STATS.endDateFormatted + ' vs ' + STATS.startDateFormattedOld + ' - ' + STATS.endDateFormattedOld + ')';
  
  STATS.statsNew = AdWordsApp.currentAccount().getStatsFor(STATS.startDate, STATS.endDate);
  STATS.statsOld = AdWordsApp.currentAccount().getStatsFor(STATS.startDateOld, STATS.endDateOld);
  
  return STATS;
}

function getStatsByWeek() {
  var STATS = {};
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = date.getDay();
  if(day == 0) { day = 7; }
  date.setDate(date.getDate()-day);
  date.setHours(12);
  
  STATS.endDate = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.endDateFormatted = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  date.setDate(date.getDate()-6);
  STATS.startDate = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.startDateFormatted = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  date.setDate(date.getDate()-1);
  STATS.endDateOld = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.endDateFormattedOld = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  date.setDate(date.getDate()-6);
  STATS.startDateOld = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  STATS.startDateFormattedOld = Utilities.formatDate(date, 'PST', 'MMM dd');
  
  STATS.dateRange = '(' + STATS.startDateFormatted + ' - ' + STATS.endDateFormatted + ' vs ' + STATS.startDateFormattedOld + ' - ' + STATS.endDateFormattedOld + ')';
  
  STATS.statsNew = AdWordsApp.currentAccount().getStatsFor(STATS.startDate, STATS.endDate);
  STATS.statsOld = AdWordsApp.currentAccount().getStatsFor(STATS.startDateOld, STATS.endDateOld);
  
  return STATS;
}

function getReportRowWeekly(sheet,accName) {
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var reportRow = -1;
  for(var k in data) {
    if(data[k][0] != accName) { continue; }
    reportRow = parseInt(k,10)+3;
  }
  
  if(reportRow != -1) { return reportRow; }
  reportRow = sheet.getLastRow()+1;
  sheet.getRange(reportRow,1,1,1).setValue(accName);
  return reportRow;
}

function getSheetByNameKey(ss, name, key) {
  var sheet = ss.getSheetByName(name);
  if(sheet) { return sheet; }
  
  sheet = ss.insertSheet(name);
  
  if(key) {
    setupReportSheet(sheet);
  }
  
  return sheet;
}

function setupReportSheet(sheet) {
  var header = ['Client name','Clicks','% Change','Impressions','% Change','Cost','% Change','Avg CPC','% Change','Conversions','% Change','CPA','% Change',
                'Clicks','% Change','Impressions','% Change','Cost','% Change','Avg CPC','% Change','Conversions','% Change','CPA','% Change'];
  
  sheet.getRange(2,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#ffe599').setBorder(true,true,true,true,true,true);
  sheet.getRange(1,2,1,12).merge();
  
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);
  
  if((sheet.getMaxColumns() - sheet.getLastColumn()) > 0) {
    sheet.deleteColumns(sheet.getLastColumn()+1, sheet.getMaxColumns() - sheet.getLastColumn());
  }
}




function cleanMasterManagementSheet(ACCOUNT_INPUTS) {
  var masterSS = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL);
  var data = masterSS.getSheetByName(MANEGEMENT_URS_TAB).getDataRange().getValues();
  data.shift();
  
  var map = {};
  for(var k in data) {
    map[data[k][0]] = 1; 
  }
  
  var data = masterSS.getSheetByName(DASHBOARD_URLS_TAB).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    map[data[k][0]] = 1; 
  }
  
  var names = {};
  for(var labelName in map) {
    var accountIter = MccApp.accounts().withCondition('LabelNames CONTAINS "'+labelName+'"').get();
    while(accountIter.hasNext()) {
      names[accountIter.next().getName()] = 1; 
    }
  }
  
  
  var sheet = SpreadsheetApp.openByUrl(MASTER_MANAGEMENT_URL).getSheetByName('Daily Report');
  var data = sheet.getDataRange().getValues();
  
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  
  for(var k = data.length-1; k>=0; k--) {
    if(!names[data[k][1]]) {
      var row = parseInt(k,10)+6;
      //Logger.log(row);
      sheet.deleteRow(row)
    }
  }
  
  var sheet = SpreadsheetApp.openByUrl(MASTER_MANAGEMENT_URL).getSheetByName('Weekly & Monthly Report');
  var data = sheet.getDataRange().getValues();
  
  data.shift();
  data.shift();
  
  for(var k = data.length-1; k>=0; k--) {
    if(!names[data[k][0]]) {
      var row = parseInt(k,10)+3;
      //Logger.log(row);
      sheet.deleteRow(row)
    }
  }
}
