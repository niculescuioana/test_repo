/***********************************************
* Tier 1 Hourly Report - MCC
* @author: Naman Jindal (nj.itprof@gmail.com)
* @version: 1.0
************************************************/

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1Z4hEsItWs3-01wF7rVfXmyr3MlS5nCZUd0K2-D1CHj4/edit#gid=0';
var REPORT_TAB = 'Report';
var GRAPH_REPORT_TAB = 'Graphs';
var GRAPH_RAW_DATA_TAB = 'Graph Data';
var LABEL = 'Tier 1';

var EMAIL_DANGER_ACCOUNTS = 'charlie@pushgroup.co.uk,naman@pushgroup.co.uk';

var DANGER_ACCOUNTS_DATE = [1,11,21];

var CPA_CHANGE_THRESHOLD_PCT = 30;
var CONVERSION_CHANGE_THRESHOLD_PCT = 30;


var FOLDER_ID = '0B51HFuINK5uhX05qUkFlY3htNnM';

/******************************************************************/

var DATE_NOW = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss');
var EMAIL_UTIL_URL = 'https://docs.google.com/spreadsheets/d/1VfKhDpiFeMFPjMBiYT5p4vvimZ-wPG1DgtsZbMVRV7I/edit#gid=0';
var EMAIL_UTIL_TAB_NAME = 'Tasks';


var CPA_CHANGE_THRESHOLD = CPA_CHANGE_THRESHOLD_PCT / 100;
var CONVERSION_CHANGE_THRESHOLD = - (CONVERSION_CHANGE_THRESHOLD_PCT/100);


function main() {
  log('Started');
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  Logger.log(now);
  
  if(DANGER_ACCOUNTS_DATE.indexOf(now.getDate()) >= 0 && now.getHours() == 11) {
    findAndReportDangerAccounts();
    return;
  }
  
  exportAdWordsStats();
  
  //log('Completed');
}


/******************** Danger Account Alert START *********************/


function findAndReportDangerAccounts() {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));  
  //var date = new Date('Apr 1, 2016');
  date.setDate(date.getDate()-1);
  date.setHours(12);
  
  var newMonthEnd = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  var newMonthEndFormatted = Utilities.formatDate(date, 'PST', 'MMM d, yyyy');
  
  var newMonthStart = newMonthEnd.substring(0,6) + '01';
  var newMonthStartFormatted = newMonthEndFormatted.split(' ')[0] + ' 1, ' + newMonthEndFormatted.split(' ')[2];
  
  var newMonthRange = newMonthStartFormatted + ' - ' + newMonthEndFormatted;
  
  var newYearEnd = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  var newYearEndFormatted = Utilities.formatDate(date, 'PST', 'MMM d, yyyy');
  
  var newYearStart = newYearEnd.substring(0,4) + '0101';
  var newYearStartFormatted = 'Jan 1, ' + newYearEndFormatted.split(' ')[2];
  
  var newYearRange = newYearStartFormatted + ' - ' + newYearEndFormatted;
  
  date.setYear(date.getYear()-1);
  
  var oldYearEnd = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  var oldYearEndFormatted = Utilities.formatDate(date, 'PST', 'MMM d, yyyy');
  
  var oldYearStart = oldYearEnd.substring(0,4) + '0101';
  var oldYearStartFormatted = 'Jan 1, ' + oldYearEndFormatted.split(' ')[2];
  
  var oldYearRange = oldYearStartFormatted + ' - ' + oldYearEndFormatted;
  
  date.setYear(date.getYear()+1);
  
  if(now.getDate() == 1) {
    date.setDate(0);
  } else {
    date.setMonth(date.getMonth()-1);
  }
  var oldMonthEnd = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  var oldMonthEndFormatted = Utilities.formatDate(date, 'PST', 'MMM d, yyyy');
  
  var oldMonthStart = oldMonthEnd.substring(0,6) + '01';
  var oldMonthStartFormatted = oldMonthEndFormatted.split(' ')[0] + ' 1, ' + oldMonthEndFormatted.split(' ')[2];
  
  var oldMonthRange = oldMonthStartFormatted + ' - ' + oldMonthEndFormatted;
  
  var YEARLY_COMPARISON_RANGE = '<b>' + newYearRange + '</b> vs <b>' + oldYearRange + '</b>';
  var MONTHLY_COMPARISON_RANGE = '<b>' + newMonthRange + '</b> vs <b>' + oldMonthRange + '</b>';  
  
  var monthlyReport = [];
  var yearlyReport = [];
  
  var mccAccount = AdWordsApp.currentAccount();
  var accounts = MccApp.accounts()
  .withCondition('Cost > 0')
  .forDateRange('LAST_30_DAYS')
  //.withLimit(5)
  .get();
  
  while(accounts.hasNext()) {
    var account = accounts.next();
    MccApp.select(account);
    
    checkAccount(newMonthStart, newMonthEnd, oldMonthStart, oldMonthEnd, monthlyReport);
    checkAccount(newYearStart, newYearEnd, oldYearStart, oldYearEnd, yearlyReport);
  }
  
  if(monthlyReport.length > 0 || yearlyReport.length > 0) {
    sendDangerReport(monthlyReport,MONTHLY_COMPARISON_RANGE,yearlyReport,YEARLY_COMPARISON_RANGE);
  }
}

function checkAccount(newStart, newEnd, oldStart, oldEnd, toAlert) {
  var report = { 'accountName': AdWordsApp.currentAccount().getName() };
  
  var statsNew = AdWordsApp.currentAccount().getStatsFor(newStart, newEnd);
  var statsOld = AdWordsApp.currentAccount().getStatsFor(oldStart, oldEnd);
  
  report.conversionsNew = statsNew.getConversions();
  report.costNew = statsNew.getCost();
  report.cpaNew = report.conversionsNew == 0 ? 0 : round((report.costNew/report.conversionsNew),2);
  
  report.conversionsNew = statsNew.getConversions();
  report.costNew = statsNew.getCost();
  report.cpaNew = report.conversionsNew == 0 ? 0 : round((report.costNew/report.conversionsNew),2);
  
  report.conversionsOld = statsOld.getConversions();
  report.costOld = statsOld.getCost();
  report.cpaOld = report.conversionsOld == 0 ? 0 : round((report.costOld/report.conversionsOld),2);
  
  report.cpaChange = report.cpaOld == 0 ? 0 : ((report.cpaNew - report.cpaOld) / report.cpaOld);
  report.conversionsChange = report.conversionsOld == 0 ? 0 : ((report.conversionsNew - report.conversionsOld) / report.conversionsOld);
  
  report.cpaAlert = 0;
  report.conversionsAlert = 0;
  
  if(report.cpaChange >= CPA_CHANGE_THRESHOLD) {
    report.cpaAlert = 1;
  }
  
  if(report.conversionsChange <= CONVERSION_CHANGE_THRESHOLD) {
    report.conversionsAlert = 1;
  }
  
  if(report.cpaAlert || report.conversionsAlert) { 
    report.cpaChange = round(100*report.cpaChange, 2)+'%';
    report.conversionsChange = round(100*report.conversionsChange, 2)+'%';
    toAlert.push(report);
  } 
}

function sendDangerReport(monthlyReport,MONTHLY_COMPARISON_RANGE,yearlyReport,YEARLY_COMPARISON_RANGE) {
  log('Sending Danger Report Email');
  var sub = "AdWords MCC Script - Danger Accounts Report";
  var msg = "";  
  
  var htmlBody = '<html><head></head><body><br><br>';    
  htmlBody += 'Hi,';
  
  if(monthlyReport.length > 0) {
    htmlBody += '<br><br>The performance of the following accounts have dropped significantly for this month as compared to last month.<br>';
    htmlBody += 'Comparison Date Range: ' + MONTHLY_COMPARISON_RANGE + '<br><br>';
    htmlBody += buildReportTable(monthlyReport, 'Month'); 
  }
  
  if(yearlyReport.length > 0) {
    htmlBody += '<br><br>The performance of the following accounts have dropped significantly for this year as compared to last year.<br>';
    htmlBody += 'Comparison Date Range: ' + YEARLY_COMPARISON_RANGE + '<br><br>';
    htmlBody += buildReportTable(yearlyReport, 'Year'); 
  }
  
  htmlBody += '<br><br>Thanks</body></html>';
  
  var EMAIL_UTIL_TAB = SpreadsheetApp.openByUrl(EMAIL_UTIL_URL).getSheetByName(EMAIL_UTIL_TAB_NAME);
  EMAIL_UTIL_TAB.appendRow([EMAIL_DANGER_ACCOUNTS, sub, msg, htmlBody, DATE_NOW]);
}


function buildReportTable(toReport,key) {
  var table = new HTMLTable();
  table.setTableStyle(['font-family: "Lucida Sans Unicode","Lucida Grande",Sans-Serif;',
                       'font-size: 12px;',
                       'background: #fff;',
                       'margin: 45px;',
                       'width: 480px;',
                       'border-collapse: collapse;',
                       'text-align: left'].join(''));
  table.setHeaderStyle(['font-size: 14px;',
                        'font-weight: normal;',
                        'color: #039;',
                        'padding: 10px 8px;',
                        'border-bottom: 2px solid #6678b1'].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  table.addHeaderColumn('Account Name');
  
  table.addHeaderColumn('Conversions Now');
  table.addHeaderColumn('Conversions Last ' + key);  
  table.addHeaderColumn('Conversions Change');
  
  table.addHeaderColumn('CPA Now');
  table.addHeaderColumn('CPA Last ' + key); 
  table.addHeaderColumn('CPA Change');
  
  for(var k in toReport) {
    table.newRow();
    table.addCell(toReport[k].accountName);
    
    table.addCell(toReport[k].conversionsNew);
    table.addCell(toReport[k].conversionsOld);
    table.addCell(toReport[k].conversionsChange);
    
    table.addCell(toReport[k].cpaNew);
    table.addCell(toReport[k].cpaOld);
    table.addCell(toReport[k].cpaChange);
  }
  
  return table.toString();
}

/****** Dange Account Alert END ****/


/************************** AdWords Reporting ************************************/

function exportAdWordsStats() {
  var FILE = 'Tier1Report.json';
  var lastCheckStats = readJSONFile(FILE);
  if(!lastCheckStats) { lastCheckStats = {}; }
  
  MccApp.accounts()
  .withCondition('LabelNames CONTAINS "'+LABEL+'"')
  .orderBy('Cost DESC').forDateRange('LAST_30_DAYS')
  .withLimit(40)
  .executeInParallel('exportAdWordsStatsInParallel', 'combineResults', JSON.stringify(lastCheckStats));
  
}

function exportAdWordsStatsInParallel(stats) {
  var accId = AdWordsApp.currentAccount().getCustomerId(); 
  var statsOld = JSON.parse(stats)[accId];
  
  var graphData = [];
  var lastHourStats = {};
  var thirtyDayAverage = {};
  
  
  var stats = AdWordsApp.currentAccount().getStatsFor('TODAY');
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var lastCheckStats = { 
    'clicks': stats.getClicks(),
    'impressions': stats.getImpressions(), 
    'conversions': stats.getConversions() 
  }
  
  if(NOW.getHours() != 0 && statsOld) {
    var thirtyDayStats = AdWordsApp.currentAccount().getStatsFor('LAST_30_DAYS');
    thirtyDayAverage = { 
      'clicks': thirtyDayStats.getClicks()/(24*30),
      'impressions': thirtyDayStats.getImpressions()/(24*30), 
      'conversions': thirtyDayStats.getConversions()/(24*30)
    }
    
    lastHourStats = { 
      'clicks': (stats.getClicks() - statsOld.clicks),
      'impressions': (stats.getImpressions() - statsOld.impressions),
      'conversions': (stats.getConversions() - statsOld.conversions)
    }
    
    getHourlyStats(graphData); 
  } else {
    lastHourStats = '';
  }
  
  
  return JSON.stringify({
    'accId': AdWordsApp.currentAccount().getCustomerId(), 'accName': AdWordsApp.currentAccount().getName(),
    'graphData': graphData, 'thirtyDayAverage': thirtyDayAverage, 
    'lastHourStats': lastHourStats, 'lastCheckStats': lastCheckStats
  });
  
}



function combineResults(results) {
  Logger.log('Processing Results');
  
  var FILE = 'Tier1Report.json';
  var lastCheckStats = readJSONFile(FILE);
  if(!lastCheckStats) { lastCheckStats = {}; }
  
  var graphData = [];
  var lastHourStats = {};
  var thirtyDayAverage = {};
  var nameMap = {};
  var ids = [];
  
  for(var z in results) {
    if(!results[z].getReturnValue()) { continue; }
    var res = JSON.parse(results[z].getReturnValue());
    ids.push(res.accId.replace(/-/g,''));
    nameMap[res.accId] = res.accName;
    if(res.lastHourStats) {
      lastHourStats[res.accId] = res.lastHourStats;
      thirtyDayAverage[res.accId] = res.thirtyDayAverage;
      graphData = graphData.concat(res.graphData);
    }
    lastCheckStats[res.accId] = res.lastCheckStats;
  }
  
  
  var mccAccount = AdWordsApp.currentAccount(); 
  
  var accountIter = MccApp.accounts()
  .withCondition('LabelNames CONTAINS "'+LABEL+'"')
  .withCondition('ExternalCustomerId NOT_IN ["' + ids.join('","') + '"]')
  .get();
  while(accountIter.hasNext()) {
    var account = accountIter.next();
    MccApp.select(account);
    var accId = AdWordsApp.currentAccount().getCustomerId();    
    var stats = AdWordsApp.currentAccount().getStatsFor('TODAY');
    
    nameMap[accId] = AdWordsApp.currentAccount().getName();
    
    var statsOld = lastCheckStats[accId];
    var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    lastCheckStats[accId] = { 
      'clicks': stats.getClicks(),
      'impressions': stats.getImpressions(), 
      'conversions': stats.getConversions() 
    }
    
    if(NOW.getHours() == 0 || !statsOld) {
      continue;
    }
    
    var thirtyDayStats = AdWordsApp.currentAccount().getStatsFor('LAST_30_DAYS');
    
    thirtyDayAverage[accId] = { 
      'clicks': thirtyDayStats.getClicks()/(24*30),
      'impressions': thirtyDayStats.getImpressions()/(24*30), 
      'conversions': thirtyDayStats.getConversions()/(24*30)
    }
    
    lastHourStats[accId] = { 
      'clicks': (stats.getClicks() - statsOld.clicks),
      'impressions': (stats.getImpressions() - statsOld.impressions),
      'conversions': (stats.getConversions() - statsOld.conversions)
    }
    
    getHourlyStats(graphData); 
  }
  
  Logger.log('Saving Results');
  writeJSONFile(FILE,lastCheckStats);  
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(REPORT_TAB);
  
  var emptyRow = ['','','',''];
  var data = [emptyRow];
  
  var accountNames = [];
  var count = 0;
  for(accId in lastHourStats) {
    accountNames.push(nameMap[accId]);
    
    data.push([nameMap[accId],'','','']);
    data.push(emptyRow);
    data.push(['','Last Hour','Average Over 30 days','Change']);
    data.push(['Clicks', lastHourStats[accId].clicks, thirtyDayAverage[accId].clicks,((lastHourStats[accId].clicks - thirtyDayAverage[accId].clicks)/thirtyDayAverage[accId].clicks)]);
    data.push(['Impressions', lastHourStats[accId].impressions, thirtyDayAverage[accId].impressions,((lastHourStats[accId].impressions - thirtyDayAverage[accId].impressions)/thirtyDayAverage[accId].impressions)]);
    data.push(['Conversions', lastHourStats[accId].conversions, thirtyDayAverage[accId].conversions,((lastHourStats[accId].conversions - thirtyDayAverage[accId].conversions)/thirtyDayAverage[accId].conversions)]);
    
    data.push(emptyRow);
    data.push(emptyRow); 
    count++;     
  } 
  
  if(count == 0) { return; }
  
  if(sheet.getLastRow() > 0) {
    sheet.getRange(1,1,sheet.getLastRow(),3).clear();
  }
  sheet.getRange(1,1,data.length,data[0].length).setValues(data);
  
  sheet.getRange(1,1,data.length,1).setFontWeight('bold');
  
  var j=0, row = 2;
  while(j < count) {
    j++;
    sheet.getRange(row,1,1,1).setBackground('#c9daf8');
    sheet.getRange(row+2,2,1,3).setFontWeight('bold').setBackground('#cfe2f3');
    sheet.getRange(row+3,1,3,1).setBackground('#fce5cd');    
    row+=8;
  }
  
  sheet.getDataRange().setFontFamily('Calibri'); 
  
  var graphDataTab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(GRAPH_RAW_DATA_TAB);
  graphDataTab.clearContents();
  
  graphDataTab.getRange(1,1,graphData.length,graphData[0].length).setValues(graphData);
  
  graphDataTab.getDataRange().setFontFamily('Calibri').setFontSize(9);
  graphDataTab.getRange(1,2,graphDataTab.getLastRow(),1).setNumberFormat('@STRING@');
  graphDataTab.getRange(1,6,graphDataTab.getLastRow(),1).setNumberFormat('@STRING@');
  graphDataTab.getRange(1,10,graphDataTab.getLastRow(),1).setNumberFormat('@STRING@');
  
  
  var graphTab = getSheet(GRAPH_REPORT_TAB);
  
  var dataRangeStartRow = 2;
  var num = 1;
  
  for(var k in accountNames) {
    graphTab.getRange(num, 1, 1, 10).merge().setValue(accountNames[k]).setBackground('#c9daf8');
    graphTab.getRange(num+2, 1, 1, 1).setValue('Clicks').setBackground('#f6b26b');
    
    insertChart(dataRangeStartRow, 2, num+3, graphTab, graphDataTab, 'Clicks', '#f6b26b');
    num += 13;
    
    graphTab.getRange(num, 1, 1, 1).setValue('Impressions').setBackground('#e06666');
    
    insertChart(dataRangeStartRow, 6, num+1, graphTab, graphDataTab, 'Impressions', '#e06666');
    num += 11;
    
    graphTab.getRange(num, 1, 1, 1).setValue('Conversions').setBackground('#93c47d');
    
    insertChart(dataRangeStartRow, 10, num+1, graphTab, graphDataTab, 'Conversions', '#93c47d');
    num += 12;
    
    dataRangeStartRow += 27;
  }
  
  graphTab.getDataRange().setFontFamily('Calibri').setFontSize(9).setFontWeight('bold');
}


function getSheet(name) {
  var ss = SpreadsheetApp.openByUrl(REPORT_URL);
  var sheet = ss.getSheetByName(name);
  if(sheet) {
    ss.deleteSheet(sheet);
  }
  
  sheet = ss.insertSheet(name);
  sheet.insertRowsAfter(900, 2500);
  return sheet;
}

function insertChart(dataRangeStartRow, col, rowNum, sheet, graphDataTab, metric, color) {
  var chart = sheet.newChart().asAreaChart()
  .setBackgroundColor("#f3f3f3")
  .setOption('useFirstColumnAsDomain', 'true')
  .setOption('title', 'Hourly '+metric+' (Last 24 Hours)')
  .setOption('vAxis.title', metric)
  .setOption('hAxis.title', 'Hour Of Day') 
  .setPointStyle(Charts.PointStyle.MEDIUM)
  .addRange(graphDataTab.getRange(dataRangeStartRow, col, 25, 2))
  .setPosition(rowNum, 1, 3, 3)
  .setColors([color])
  .setOption('width', 880)
  .setOption('height', 190)
  .build();
  
  
  sheet.insertChart(chart);   
}

function getHourlyStats(graphData) {
  var dateFrom = getAdWordsFormattedDate(1,'yyyyMMdd');
  var dateTo = getAdWordsFormattedDate(0, 'yyyyMMdd');
  
  var formattedToday = getAdWordsFormattedDate(0, 'MMM dd, yyyy');
  var formattedYesterday = getAdWordsFormattedDate(1, 'MMM dd, yyyy');
  
  var today = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'),10);
  
  var emptyRow = ['','','','','','','','','','',''];
  graphData.push([AdWordsApp.currentAccount().getName(),'','','','','','','','','','']);
  graphData.push(['Date','Hour','Clicks','','Date','Hour','Impressions','','Date','Hour','Conversions']);
  
  var map = {};			 
  map[formattedYesterday] = {};
  map[formattedToday] = {};
  
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['HourOfDay', 'Date', 'Impressions', 'Clicks', 'Conversions'];
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'during',dateFrom + ',' + dateTo].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.Date != today && parseInt(row.HourOfDay,10) < hour) {
      continue;
    }
    if(row.Date == today) {
      row.Date = formattedToday; 
    } else {
      row.Date = formattedYesterday; 
    }
    
    map[row.Date][row.HourOfDay] = row;
  }
  
  for(var k = hour; k <= 23; k++) {
    var row = map[formattedYesterday][k];
    if(!row) {
      row = { HourOfDay: k, Clicks: 0, Impressions: 0, Conversions: 0 };
    }
    graphData.push([formattedYesterday, row.HourOfDay, row.Clicks, '', 
                    formattedYesterday, row.HourOfDay, row.Impressions, '',
                    formattedYesterday, row.HourOfDay, row.Conversions]);
    
  }
  
  for(var k = 0; k < hour; k++) {
    var row = map[formattedToday][k];
    if(!row) {
      row = { HourOfDay: k, Clicks: 0, Impressions: 0, Conversions: 0 };
    }
    graphData.push([formattedToday, row.HourOfDay, row.Clicks, '', 
                    formattedToday, row.HourOfDay, row.Impressions, '',
                    formattedToday, row.HourOfDay, row.Conversions]);
    
  }
  
  graphData.push(emptyRow);
}

/**************************** Utility Methods ***************************************/

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
} 

function getLastDay(format) {
  var date = new Date();
  date.setDate(0);  
  date.setHours(10);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

/*********************************************
* HTMLTable: A class for building HTML Tables
* Version 1.0
**********************************************/
function HTMLTable() {
  this.headers = [];
  this.columnStyle = {};
  this.body = [];
  this.currentRow = 0;
  this.tableStyle;
  this.headerStyle;
  this.cellStyle;
  
  this.addHeaderColumn = function(text) {
    this.headers.push(text);
  };
  
  this.addCell = function(text,style) {
    if(!this.body[this.currentRow]) {
      this.body[this.currentRow] = [];
    }
    this.body[this.currentRow].push({ val:text, style:(style) ? style : '' });
  };
  
  this.newRow = function() {
    if(this.body != []) {
      this.currentRow++;
    }
  };
  
  this.getRowCount = function() {
    return this.currentRow;
  };
  
  this.setTableStyle = function(css) {
    this.tableStyle = css;
  };
  
  this.setHeaderStyle = function(css) {
    this.headerStyle = css; 
  };
  
  this.setCellStyle = function(css) {
    this.cellStyle = css;
    if(css[css.length-1] !== ';') {
      this.cellStyle += ';';
    }
  };
  
  this.toString = function() {
    var retVal = '<table ';
    if(this.tableStyle) {
      retVal += 'style="'+this.tableStyle+'"';
    }
    retVal += '>'+_getTableHead(this)+_getTableBody(this)+'</table>';
    return retVal;
  };
  
  function _getTableHead(instance) {
    var headerRow = '';
    for(var i in instance.headers) {
      headerRow += _th(instance,instance.headers[i]);
    }
    return '<thead><tr>'+headerRow+'</tr></thead>';
  };
  
  function _getTableBody(instance) {
    var retVal = '<tbody>';
    for(var r in instance.body) {
      var rowHtml = '<tr>';
      for(var c in instance.body[r]) {
        rowHtml += _td(instance,instance.body[r][c]);
      }
      rowHtml += '</tr>';
      retVal += rowHtml;
    }
    retVal += '</tbody>';
    return retVal;
  };
  
  function _th(instance,val) {
    var retVal = '<th scope="col" ';
    if(instance.headerStyle) {
      retVal += 'style="'+instance.headerStyle+'"';
    }
    retVal += '>'+val+'</th>';
    return retVal;
  };
  
  function _td(instance,cell) {
    var retVal = '<td ';
    if(instance.cellStyle || cell.style) {
      retVal += 'style="';
      if(instance.cellStyle) {
        retVal += instance.cellStyle;
      }
      if(cell.style) {
        retVal += cell.style;
      }
      retVal += '"';
    }
    retVal += '>'+cell.val+'</td>';
    return retVal;
  };
}

function log(msg) {
  var time = Utilities.formatDate(new Date(),AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  Logger.log(time + ' - ' + msg);
}

function writeJSONFile(fileName,toWrite) {
  var file = getFile(fileName);
  try {
    file.setContent(JSON.stringify(toWrite));
  } catch(e) {
    Logger.log('Cannot store data. File exceeding max size');
  }
}

function readJSONFile(fileName) {
  var file = getFile(fileName);
  var fileData = file.getBlob().getDataAsString();
  if(fileData && fileData != 'undefined') {
    return JSON.parse(fileData);
  } else {
    return null;
  }
}


function getFile(fileName) {
  
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var fileIter = folder.getFilesByName(fileName);
  
  if(!fileIter.hasNext()) {
    var file = folder.createFile(fileName,'');
    return file;
  } else {
    return fileIter.next();
  }
}

function round(num, n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}


function readInputsForAccounts() {
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var URL_INFO_TAB = 'Dashboard Urls';

  var SETTINGS = {};
  
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(URL_INFO_TAB).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][1]) { continue; }
    var sheet = SpreadsheetApp.openByUrl(data[k][1]).getSheetByName('Common Inputs');
    if(!sheet) {
      continue;
    }
    
    var inputData = sheet.getDataRange().getValues();
    inputData.shift(); 
    var header = inputData.shift()
    
    for(var j in inputData) {
      if(SETTINGS[inputData[j][0]] && SETTINGS[inputData[j][0]].MONTHLY_BUDGET && SETTINGS[inputData[j][0]].DAILY_BUDGET) { continue; }
      SETTINGS[inputData[j][0]] = {};
      for(var l in header) {
        SETTINGS[inputData[j][0]][header[l]] = inputData[j][l];
      }
    }
  }
  
  return SETTINGS;
}