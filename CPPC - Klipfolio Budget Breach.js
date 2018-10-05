
var KLIPFOLIO_URL = 'https://docs.google.com/spreadsheets/d/1UeGv6NOl4LDBZU6H0-RveihxFZgPDyXKOXss-ltazGU/edit?ts=57e83723';

var KRATU_URL = 'https://docs.google.com/spreadsheets/d/12wsp5DYyRgVANU_2pQ5xLd4CLnth3WtGJdp6jS-aao8/edit#gid=301484764';
var KRATU_TAB_NAME = 'Budget Breach Report';

var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';

function main() {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  //Logger.log(now);
  
  var MANAGEMENT_URLS = {};
  var INPUT_SETTINGS = readInputsForAccounts(MANAGEMENT_URLS);
  
  //sendZeroClicksAlert(INPUT_SETTINGS); 
  if(now.getHours() == 8 || now.getHours() == 9) {
    compileBudgetBreachReport(MANAGEMENT_URLS, INPUT_SETTINGS); 
  }
  
  if(now.getHours() == 7) {
    sendZeroClicksAlert(INPUT_SETTINGS); 
  }
  
  //compileBudgetBreachReport(); 
  
  compileCPAReport();
}

function compileCPAReport() {
  var report = [['Account', 'CPA', 'Yesterday', 'Flag', 'Number Yesterday', 'Average']];
  var iter = MccApp.accounts().withCondition('LabelNames CONTAINS "Tier 1"').get();
  while(iter.hasNext()) {
    var acc = iter.next();
    MccApp.select(acc);
    var accName = AdWordsApp.currentAccount().getName(); 
    
    var stats_30 = AdWordsApp.currentAccount().getStatsFor('LAST_30_DAYS');
    var cost_30 = stats_30.getCost();
    var conversions_30 = stats_30.getConversions();
    var cpa_30 = conversions_30 == 0 ? 0 : round(cost_30/conversions_30,2);
    
    var stats = AdWordsApp.currentAccount().getStatsFor('YESTERDAY');
    var cost = stats.getCost();
    var conversions = stats.getConversions();
    var cpa = conversions == 0 ? 0 : round(cost/conversions,2);
    
    var variance = Math.abs(round(100*(cpa - cpa_30)/cpa_30, 2));
    report.push([accName, cpa_30, cpa, variance > 20 ? cpa : '', conversions, round(conversions_30/30,2)]);
  }
  
  var tab = SpreadsheetApp.openByUrl(KLIPFOLIO_URL).getSheetByName('CPA Report');
  tab.clearContents();
  tab.getRange(1,1,report.length,report[0].length).setValues(report);
}

function sendZeroClicksAlert(INPUT_SETTINGS) {
  var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'))
  now.getDate(now.getDate()-1);
  var day = now.getDay();
  
  var urlData = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName('Dashboard Urls').getDataRange().getValues();
  urlData.shift();
  
  var accounts = [];
  for(var j in urlData) {
    if(!urlData[j][0]) { continue; } 
    
    var iter = MccApp.accounts()
    .withCondition('LabelNames CONTAINS "' + urlData[j][0] + '"')
    .withCondition('Clicks = 0')
    .forDateRange('YESTERDAY')
    .get();
    
    while(iter.hasNext()) {
      var accName = iter.next().getName();
      if(INPUT_SETTINGS[accName] && INPUT_SETTINGS[accName].WEEKEND_OFF == 'Y' && (day == 0 || day == 1)) {
        continue;
      }
      accounts.push([accName]); 
    }
  }
  
  var tab = SpreadsheetApp.openByUrl(KLIPFOLIO_URL).getSheetByName('Zero Clicks Yesterday');
  tab.clearContents();
  
  tab.getRange(2,1,accounts.length,1).setValues(accounts);
}

function compileBudgetBreachReport(MANAGEMENT_URLS, INPUT_SETTINGS) {
  var BUDGET_BREACH_MAP = readBudgetBreachReport();
  
  var accountNames = Object.keys(BUDGET_BREACH_MAP);
  if(!accountNames.length) { return; }
  
  var reportByManager = {};
  var report = [['Account', 'Time', 'Cap', 'Spend', 'Daily Target', 'Breach last 5 days']];
  
  for(var manager in MANAGEMENT_URLS) {
    reportByManager[manager] = [report[0]];
  }
  
  var iter = MccApp.accounts()
  .withCondition('Name IN ["' + accountNames.join('","') + '"]')
  .get();
  
  while(iter.hasNext()) {
    var acc = iter.next();
    MccApp.select(acc);
    
    var accName = AdWordsApp.currentAccount().getName();
    if(!INPUT_SETTINGS[accName]) {
      Logger.log(acc.getName());
      continue;
    }
    
    var dailyTarget = '';
    if(INPUT_SETTINGS[accName].MONTHLY_BUDGET) {
      var DAYS_LEFT = getDaysLeftNew(INPUT_SETTINGS[accName]);
      dailyTarget = round((INPUT_SETTINGS[accName].MONTHLY_BUDGET - AdWordsApp.currentAccount().getStatsFor('THIS_MONTH').getCost()) / DAYS_LEFT,2);
    }
    
    report.push([accName, BUDGET_BREACH_MAP[accName].TIME, INPUT_SETTINGS[accName].DAILY_BUDGET,
                 AdWordsApp.currentAccount().getStatsFor('YESTERDAY').getCost(), 
                 dailyTarget, BUDGET_BREACH_MAP[accName].COUNT]); 
    
    var labelIter = acc.labels().get();
    while(labelIter.hasNext()) {
      var lbl = labelIter.next().getName();
      if(MANAGEMENT_URLS[lbl]) {
        reportByManager[lbl].push([accName, BUDGET_BREACH_MAP[accName].TIME, INPUT_SETTINGS[accName].DAILY_BUDGET,
                                   AdWordsApp.currentAccount().getStatsFor('YESTERDAY').getCost(), 
                                   dailyTarget, BUDGET_BREACH_MAP[accName].COUNT]);
      }
    }
  }
  
  
  var tab = SpreadsheetApp.openByUrl(KLIPFOLIO_URL).getSheetByName('Budget Breach');
  tab.clearContents();
  tab.getRange(1,1,report.length,report[0].length).setValues(report);
  tab.sort(6, false);
  SpreadsheetApp.flush();
  
  for(var manager in reportByManager) {
    var tab = SpreadsheetApp.openByUrl(MANAGEMENT_URLS[manager]).getSheetByName('Budget Breach Report');
    if(!tab) { continue; }
    tab.clearContents();
    tab.getRange(1,1,reportByManager[manager].length,reportByManager[manager][0].length).setValues(reportByManager[manager]);
    tab.sort(2);
  }
}

function readBudgetBreachReport() {
  var data = SpreadsheetApp.openByUrl(KRATU_URL).getSheetByName(KRATU_TAB_NAME).getDataRange().getValues();
  var header = data.shift();
  
  var yest = getAdWordsFormattedDate(1, 'd MMM');
  
  var SETTINGS = {};
  var index = header.indexOf(yest);
  for(var x in data) {
    if(!data[x][0] || !data[x][index]) { continue; }
    SETTINGS[data[x][0]] = { 'TIME': data[x][index], 'COUNT': 0 };
    
    var counter = 0, y = parseInt(index,10);
    while(1) {
      counter++;
      if(data[x][y]) {
        SETTINGS[data[x][0]].COUNT++;
      }
      
      y--;
      
      if(counter == 5) { break; }
    }
  }
  
  return SETTINGS;
}

function readInputsForAccounts(MANAGEMENT_URLS) {
  
  var SETTINGS = {};
  
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName('Dashboard Urls').getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][1]) { continue; }
    MANAGEMENT_URLS[data[k][0]] = data[k][2];
    var sheet = SpreadsheetApp.openByUrl(data[k][1]).getSheetByName('Account Inputs');
    if(!sheet) {
      continue;
    }
    
    var inputData = sheet.getDataRange().getValues();
    inputData.shift(); 
    var header = inputData.shift()
    
    for(var j in inputData) {
      if(SETTINGS[inputData[j][0]] && SETTINGS[inputData[j][0]].MONTHLY_BUDGET /*&& SETTINGS[inputData[j][0]].DAILY_BUDGET*/) { continue; }
      SETTINGS[inputData[j][0]] = {};
      for(var l in header) {
        SETTINGS[inputData[j][0]][header[l]] = inputData[j][l];
      }
    }
  }
  
  return SETTINGS;
}	


function getDaysLeft() {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var hour = now.getHours();
  var mm = now.getMonth()+1;
  var yyyy = now.getYear();
  
  var monthEnd = new Date(yyyy, mm, 0);
  monthEnd.setHours(22);
  
  var daysInMonth = monthEnd.getDate();
  var daysPassed = now.getDate() - 1 + hour/24;
  
  return round(daysInMonth - daysPassed,2);
}

function getDaysLeftNew(ACCOUNT_INPUTS) {
  var BUSINESS_DAYS = 7;
  if(ACCOUNT_INPUTS.WEEKEND_OFF == 'Y') {
    BUSINESS_DAYS = 5;
  }  
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  
  var hour = now.getHours();
  var mm = now.getMonth()+1;
  var yyyy = now.getYear();
  
  var days = ['January','February','March','April','May','June','July',
              'August','September','October','November','December'];
  var monthName = days[now.getMonth()];
  
  var date = now.getDate();
  var monthEnd = new Date(yyyy, mm, 0);
  monthEnd.setHours(22);
  
  var DAYS_IN_MONTH = monthEnd.getDate();
  var DAYS_LEFT = DAYS_IN_MONTH - now.getDate();
  
  if(BUSINESS_DAYS == 5) {
    DAYS_IN_MONTH = getWeekdaysInMonth(mm-1, yyyy, DAYS_IN_MONTH);
    DAYS_LEFT = getWorkingDaysNew(now, monthEnd) - 1;
  }
  
  
  return (DAYS_LEFT + 1 - hour/24);  
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}



function getWeekdaysInMonth(month, year, days) {
  var weekdays = 0;
  for(var i=0; i< days; i++) {
    if (isWeekday(year, month, i+1)) weekdays++;
  }
  return weekdays;
}

function isWeekday(year, month, day) {
  var day = new Date(year, month, day).getDay();
  return day !=0 && day !=6;
}

function getWorkingDaysNew(startDate, endDate){
  var weekDays = 0;
  
  var currentDate = new Date(startDate.valueOf());
  while (currentDate <= endDate)  {  
    var day = currentDate.getDay();
    if(day != 0 && day != 6) { weekDays++; }
    currentDate.setDate(currentDate.getDate()+1); 
  }
  
  if(currentDate > endDate && currentDate.getDate() == endDate.getDate()) {
    var day = currentDate.getDay();
    if(day != 0 && day != 6) { weekDays++; }
  }
  
  return weekDays;
}