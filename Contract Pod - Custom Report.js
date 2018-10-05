/*************************************************
* Weekly and Monthly Reporting
* @version: 1.0
* @author: Naman Jindal (nj.itprof@gmail.com)
***************************************************/

var WEEKLY_TAB_NAME = 'Weekly Performance';
var MONTHLY_TAB_NAME = 'Monthly Performance';


var START_DATE = '20171001';
var LAST_N_MONTHS = 3;

function main() {
  var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1S9VmTuqyjQd4O7b0gaSKhy7Ew0p2UGBK_GVhO2hNnvY/edit';
  runScript("Contract Pod", REPORT_URL);
}

function runScript(ACC_NAME, REPORT_URL) {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var PROFILE_ID = 94401520;
  
  var mccAccount = AdWordsApp.currentAccount();
  MccApp.select(MccApp.accounts().withCondition('Name = "'+ ACC_NAME +'"').get().next());
  
  compileMonthlyReport(REPORT_URL, PROFILE_ID, 0);
  if(now.getDate() < 3) {
    compileMonthlyReport(REPORT_URL, PROFILE_ID, 1); 
  }
 
  compileWeeklyReport(REPORT_URL, PROFILE_ID, 0); 
  if(now.getDay() > 0 && now.getDay() < 3) {
    compileWeeklyReport(REPORT_URL, PROFILE_ID, 1);
  }

  compileWeeklyReportByCountry(REPORT_URL, PROFILE_ID, 0); 
  if(now.getDay() > 0 && now.getDay() < 3) {
    compileWeeklyReportByCountry(REPORT_URL, PROFILE_ID, 1);
  }
  
  compileSourceReport(REPORT_URL, PROFILE_ID);
}

function compileSourceReport(REPORT_URL, PROFILE_ID) {
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Analytics Top traffic Sources');
  var stats = {};
  
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:pageViews,ga:sessions,ga:goalCompletionsAll,ga:bounceRate,ga:pageviewsPerSession,ga:avgSessionDuration",
        {'dimensions': 'ga:sourceMedium', 'sort': '-ga:pageViews'});
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var z in rows) {
    var sm = rows[z].shift();
    if(!stats[sm]) {
      stats[sm] = {
        'lm': [0, 0, 0, 0, 0, 0],
        'mtd': [0, 0, 0, 0, 0, 0],
        'lmtd': [0, 0, 0, 0, 0, 0]
      }
    }
    
    stats[sm]['mtd'] = rows[z];
  }
  
  var dt = new Date(getAdWordsFormattedDate(1, 'yyyy-MM-dd'));
  dt.setHours(12);
  dt.setMonth(dt.getMonth()-1);
  var TO = Utilities.formatDate(dt, 'PST', 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:pageViews,ga:sessions,ga:goalCompletionsAll,ga:bounceRate,ga:pageviewsPerSession,ga:avgSessionDuration",
        {'dimensions': 'ga:sourceMedium', 'sort': '-ga:pageViews'});
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var z in rows) {
    var sm = rows[z].shift();
    if(!stats[sm]) {
      stats[sm] = {
        'lm': [0, 0, 0, 0, 0, 0],
        'mtd': [0, 0, 0, 0, 0, 0],
        'lmtd': [0, 0, 0, 0, 0, 0]
      }
    }
    
    stats[sm]['lmtd'] = rows[z];
  }
  
  var dt = new Date(getAdWordsFormattedDate(1, 'yyyy-MM-dd'));
  dt.setHours(12);
  dt.setDate(1);
  dt.setDate(0);
  var TO = Utilities.formatDate(dt, 'PST', 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:pageViews,ga:sessions,ga:goalCompletionsAll,ga:bounceRate,ga:pageviewsPerSession,ga:avgSessionDuration",
        {'dimensions': 'ga:sourceMedium', 'sort': '-ga:pageViews'});
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var z in rows) {
    var sm = rows[z].shift();
    if(!stats[sm]) {
      stats[sm] = {
        'lm': [0, 0, 0, 0, 0, 0],
        'mtd': [0, 0, 0, 0, 0, 0],
        'lmtd': [0, 0, 0, 0, 0, 0]
      }
    }
    
    stats[sm]['lm'] = rows[z];
  }
  
  var rows = [];
  for(var sm in stats) {
    var row = stats[sm];
    var changeRow = [];
    for(var z in row['mtd']) {
      changeRow.push(row['lmtd'][z] == 0 ? 0 : (row['mtd'][z] - row['lmtd'][z]) / row['lmtd'][z]); 
    }
    
    rows.push([sm].concat(row['lm'].concat(row['mtd'].concat(changeRow)))); 
  }
  
  tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  tab.getRange(3,1,rows.length,rows[0].length).setValues(rows);
}

function compileWeeklyReportByCountry(REPORT_URL, PROFILE_ID, factor) {
  
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day - 1 + 7*factor;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(sd, 'PST', 'yyyy-MM-dd');
  var startAdWords = startDate.replace(/-/g, '');
  
  diff -= 6;
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  ed.setDate(ed.getDate() - diff);
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endDate = Utilities.formatDate(ed, 'PST', 'yyyy-MM-dd');
  var endAdWords = endDate.replace(/-/g, '');  

  var key = start;
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CountryCriteriaId', 'Impressions','Clicks','AverageCpc','Conversions','Cost','Ctr',
              'AveragePosition','CostPerConversion','ConversionRate','AllConversions'];
  var report = 'GEO_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'during',startAdWords+','+endAdWords].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  
  var map = {};
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row['CountryCriteriaId'] != 'United States' && row['CountryCriteriaId'] != 'United Kingdom') { continue; }
    
    var rowKey = key + '  (' + row['CountryCriteriaId'] + ')';
    map[rowKey] = {};
    for(var k in cols) {
      map[rowKey][cols[k]] = row[cols[k]];
    }
  }
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Country Data');
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var found = false;
  var rowNum = -1;
  
  for(var rowKey in map) {
    for(var k in data) {
      if(data[k][0] == rowKey) {
        rowNum = parseInt(k,10) + 3;
        found = true;
        break;
      }
    }
    
    if(!found) {
      rowNum = 3;
      sheet.insertRowBefore(3);
    } 
  
    var row = [rowKey, map[rowKey].Clicks, map[rowKey].Impressions, map[rowKey].Ctr, map[rowKey].AverageCpc, map[rowKey].Cost, map[rowKey].AveragePosition, 
               map[rowKey].Conversions, map[rowKey].CostPerConversion, map[rowKey].ConversionRate, map[rowKey].AllConversions];
  
    sheet.getRange(rowNum,1,1,row.length).setValues([row]);
  }
}

function compileWeeklyReport(REPORT_URL, PROFILE_ID, factor) {
  //initialWeeklySetup();
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day - 1 + 7*factor;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(sd, 'PST', 'yyyy-MM-dd');
  var startAdWords = startDate.replace(/-/g, '');
  
  diff -= 6;
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  ed.setDate(ed.getDate() - diff);
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endDate = Utilities.formatDate(ed, 'PST', 'yyyy-MM-dd');
  var endAdWords = endDate.replace(/-/g, '');  

  var key = start + ' - ' + end;  
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['Impressions','Clicks','AverageCpc','Conversions','Cost','Ctr',
              'AveragePosition','CostPerConversion','ConversionRate','AllConversions'];
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'during',startAdWords+','+endAdWords].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  
  var map = {};
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    for(var k in cols) {
      map[cols[k]] = row[cols[k]];
    }
  }
  
  var analyticsStats = getDataFromAnalytics(PROFILE_ID, startDate, endDate);
  
  var stats = AdWordsApp.currentAccount().getStatsFor(startAdWords,endAdWords);
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(WEEKLY_TAB_NAME);
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var found = false;
  var rowNum = -1;
  var cpa = stats.getConversions() == 0 ? 0 : stats.getCost() / stats.getConversions();
  for(var k in data) {
    if(data[k][0].split(' - ')[0] == key.split(' - ')[0]) {
      rowNum = parseInt(k,10) + 3;
      found = true;
      break;
    }
  }
  
  if(!found) {
    rowNum = 3;
    sheet.insertRowBefore(3);
    sheet.getRange(4,14,1,5).copyTo(sheet.getRange(3,14,1,5));
  } 
  
  var row = [key, map.Clicks, map.Impressions, map.Ctr, map.AverageCpc, map.Cost, map.AveragePosition, 
             map.Conversions, map.CostPerConversion, map.ConversionRate, map.AllConversions];
  
  analyticsStats['PCT'] = analyticsStats['PPCSessions'] / analyticsStats['Sessions'];
  for(var metric in analyticsStats) {
    row.push(analyticsStats[metric]);
  }
  
  sheet.getRange(rowNum,1,1,row.length).setValues([row]);
}

function compileMonthlyReport(REPORT_URL, PROFILE_ID, factor) {
  var keys = [];
  var date  = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  date.setHours(12);
  
  if(factor > 0){
    date.setDate(1);
  }
  
  while(factor > 0) {
    date.setDate(0);
    factor--;
  }
  
  var key = Utilities.formatDate(date, AdWordsApp.currentAccount().getTimeZone(), 'MMMM yyyy');
  var endDate = Utilities.formatDate(date, AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  var endAdWords = endDate.replace(/-/g, '');  
  
  var startDate = endDate.substring(0,8) + '01';
  var startAdWords = startDate.replace(/-/g, '');  
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['Impressions','Clicks','AverageCpc','Conversions','Cost','Ctr',
              'AveragePosition','CostPerConversion','ConversionRate','AllConversions'];
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'during',startAdWords+','+endAdWords].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  
  var map = {};
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    for(var k in cols) {
      map[cols[k]] = row[cols[k]];
    }
  }
  
  var analyticsStats = getDataFromAnalytics(PROFILE_ID, startDate, endDate);
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(MONTHLY_TAB_NAME);
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var found = false;
  var rowNum = -1;
  
  for(var k in data) {
    if(data[k][0] == key) {
      rowNum = parseInt(k,10) + 3;
      found = true;
      break;
    }
  }
  
  if(!found) {
    rowNum = 3;
    sheet.insertRowBefore(3);
  } 
  
  var row = [key, map.Clicks, map.Impressions, map.Ctr, map.AverageCpc, map.Cost, map.AveragePosition, 
             map.Conversions, map.CostPerConversion, map.ConversionRate, map.AllConversions];
  
  analyticsStats['PCT'] = analyticsStats['PPCSessions'] / analyticsStats['Sessions'];
  for(var metric in analyticsStats) {
    row.push(analyticsStats[metric]);
  }
  
  sheet.getRange(rowNum,1,1,row.length).setValues([row]);
}


function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function info(msg) {
  var time = Utilities.formatDate(new Date(),AdWordsApp.currentAccount().getTimeZone(),'MMM dd, yyyy HH:mm:ss.SSS');
  Logger.log(time + ' - ' + msg);
}  

function getDataFromAnalytics(PROFILE_ID,FROM,TO) {
  var attempts = 3;
  var stats = {
    'Sessions': 0, 'PPCSessions': 0, 'PCT': 0, 'Goals': 0, 'PPCGoals': 0, 
    'PPCCalls': 0, 'PPCCallsAw': 0, 
    'ContactForms': 0, 'ContactFormsAw': 0, 
    'DemoRequests': 0, 'DemoRequestsAw': 0, 
    'Overlays': 0, 'OverlaysAw': 0,
    'ExitOverlays': 0, 'ExitOverlaysAw': 0
  };
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:sessions,ga:goalCompletionsAll,ga:goal1Completions,ga:goal4Completions,ga:goal9Completions,ga:goal7Completions,ga:goal3Completions",
        {'filters': 'ga:source==google;ga:medium==cpc'});
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    stats.PPCSessions = parseFloat(rows[k][0]);
    stats.PPCGoals = parseInt(rows[k][1],10);
    stats.PPCCallsAw = parseFloat(rows[k][2]);
    stats.DemoRequestsAw = parseInt(rows[k][3],10);  
    stats.OverlaysAw = parseInt(rows[k][4], 10);
    stats.ExitOverlaysAw = parseInt(rows[k][5], 10);
    stats.ContactFormsAw = parseInt(rows[k][6], 10);
  }
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:sessions,ga:goalCompletionsAll,ga:goal1Completions,ga:goal4Completions,ga:goal9Completions,ga:goal7Completions,ga:goal3Completions");
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    stats.Sessions = parseFloat(rows[k][0]);
    stats.Goals = parseFloat(rows[k][1]);
    stats.PPCCalls = parseFloat(rows[k][2]);
    stats.DemoRequests = parseInt(rows[k][3],10);  
    stats.Overlays = parseInt(rows[k][4], 10);
    stats.ExitOverlays = parseInt(rows[k][5], 10);
    stats.ContactForms = parseInt(rows[k][6], 10);
  }
  
  return stats;
}