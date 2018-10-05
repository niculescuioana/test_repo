var SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/10prWRt88JgSacSe8oXeg5kjBg5u769VRtV21RaPj750/edit';
var SETTINGS_TAB_NAME = 'Performance & Tracking';

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1fTKm5pEF5zMgHj-ivyOVg8QB8sCgYmHeVbH6FPRvK1A/edit';

function main() {
  var map = parseInputs();
  
  
  for(var id in map['an']) {
    runForAnalyticsOnly(map['an'][id]); 
  }
  
  
  var ids = Object.keys(map['aw']);
  MccApp.accounts().withIds(ids).executeInParallel('run', 'compile', JSON.stringify(map));
}

function parseInputs() {
  var map = { 'aw': {}, 'an': {} }; 
  
  var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(SETTINGS_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var header = data.shift();
  var idRow = data.shift();
  for(var k in header) {
    var idx = parseInt(k,10)+1;
    if(k == 0 || !header[k]) { continue; }
    if(!idRow[k] && !idRow[idx]) { continue; }
    
    var row = { 'ADWORDS_ID': idRow[k], 'PROFILE_ID': idRow[idx], 'NAME': header[k], 'GOALS': [] };
    for(var z in data) {
      if(!data[z][idx]) { continue; }
      
      var goal = { 'name': data[z][k], 'id': parseInt(data[z][idx].toLowerCase().split('/')[0].trim().replace('goal id ', ''), 10) };
      row['GOALS'].push(goal);
    }
    
    if(idRow[k]) {
      map['aw'][idRow[k]] = row;
    } else {
      map['an'][idRow[k]] = row;
    }
  }
  
  return map;
}

function compile() {
  
};

function runForAnalyticsOnly(SETTINGS) {
  compileThisWeekReport(SETTINGS);
  compileLastWeekReport(SETTINGS);
}

function run(input) {
  var SETTINGS = JSON.parse(input)['aw'][AdWordsApp.currentAccount().getCustomerId()];
  
  compileLastWeekReport(SETTINGS);
  compileLastMonthReport(SETTINGS);
  compileThisWeekReport(SETTINGS);
}

function compileLastMonthReport(SETTINGS) {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  //if(now.getDate() > 7) { return; }
  now.setDate(1);
  now.setDate(0);
  now.setHours(12);
  
  var end = Utilities.formatDate(now, 'PST', 'MMM d, yyyy');
  var endDate =  Utilities.formatDate(now, 'PST', 'yyyyMMdd');
  
  now.setDate(1);
  var start = Utilities.formatDate(now, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(now, 'PST', 'yyyyMMdd');
  
  var key = Utilities.formatDate(now, 'PST', 'MMM yyyy');
  compileReport(SETTINGS, key, startDate, endDate);
}

function compileThisWeekReport(SETTINGS) {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day - 1;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(sd, 'PST', 'yyyyMMdd');
  
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endDate =  Utilities.formatDate(ed, 'PST', 'yyyyMMdd');
  
  var key = start + ' - ' + end;
  compileReport(SETTINGS, key, startDate, endDate);
}

function compileLastWeekReport(SETTINGS) {
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day + 6 ;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startAdWords = Utilities.formatDate(sd, 'PST', 'yyyMMdd');
  
  
  var diff = day;
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  ed.setDate(ed.getDate() - diff);
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endAdWords = Utilities.formatDate(ed, 'PST', 'yyyMMdd');
  
  var key = start + ' - ' + end;
  compileReport(SETTINGS, key, startAdWords, endAdWords);
}


function compileReport(SETTINGS, key, startAdWords, endAdWords) {
  var ss = SpreadsheetApp.openByUrl(REPORT_URL);
  var sheet = ss.getSheetByName(SETTINGS.NAME);
  if(!sheet) {
    ss.setActiveSheet(ss.getSheetByName('Template'));
    sheet = ss.duplicateActiveSheet();
    sheet.setName(SETTINGS.NAME);
    sheet.showSheet();
    ss.getSheetByName('Template').hideSheet();
  }
  
  sheet.getRange('A:A').setNumberFormat('@STRING@');
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var found = false;
  var rowNum = -1;
  
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
  } 
  
  var row = [
    key, 0, 0, 0,
    0, 0, 0, 0, 0
  ];
  
  if(SETTINGS.ADWORDS_ID) {
    
    try {
      var stats = AdWordsApp.currentAccount().getStatsFor(startAdWords,endAdWords);
      var cpa = stats.getConversions() == 0 ? 0 : stats.getCost() / stats.getConversions();
      
      row = [
        key, stats.getClicks(), stats.getImpressions(), stats.getCtr(),
        stats.getAverageCpc(), stats.getCost(), stats.getConversions(), 
        cpa, stats.getConversionRate()
      ];
    } catch(ex) {
      
    }
  }
  
  var initMap = {
    'Sessions': 0
  };
  
  var goals = [], goalHeader = [];
  for(var z in SETTINGS.GOALS) {
    initMap[SETTINGS.GOALS[z]['name']] = 0;
    goals.push('ga:goal'+SETTINGS.GOALS[z]['id']+'Completions');
    goalHeader.push(SETTINGS.GOALS[z]['name'] + ' (All)', SETTINGS.GOALS[z]['name'] + ' (AdWords)');
  }
  
  var googleStats = JSON.parse(JSON.stringify(initMap));
  var overallStats = JSON.parse(JSON.stringify(initMap));
  
  if(SETTINGS.PROFILE_ID) {
    var FROM = startAdWords.substring(0,4) + '-' + startAdWords.substring(4,6) + '-' + startAdWords.substring(6,8);
    var TO = endAdWords.substring(0,4) + '-' + endAdWords.substring(4,6) + '-' + endAdWords.substring(6,8);
    
    var optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
    getDataFromAnalytics(SETTINGS.PROFILE_ID,googleStats,goals,FROM,TO,optArgs);
    
    getDataFromAnalytics(SETTINGS.PROFILE_ID,overallStats,goals,FROM,TO,{});
  }
  
  row.push(overallStats['Sessions'], googleStats['Sessions'], overallStats['Sessions'] == 0 ? 0 : round(googleStats['Sessions'] / overallStats['Sessions'], 4));
  for(var key in overallStats) {
    if(key == 'Sessions') { continue; }
    row.push(overallStats[key], googleStats[key]);
  }
  
  sheet.getRange(rowNum,1,1,row.length).setValues([row]);
  
  if(goalHeader.length) {
    sheet.getRange(2,13,1,goalHeader.length).setValues([goalHeader]);
  }
}



function getDataFromAnalytics(PROFILE_ID,stats,goals,FROM,TO,optArgs) {
  var attempts = 3;
  
  var metrics = "ga:sessions";
  if(goals.length) {
    metrics += ',' + goals.join(',');
  }
  
  var keys = Object.keys(stats);
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        metrics,
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    for(var z in keys) {
      stats[keys[z]] += parseInt(rows[k][z],10);
    }
  }
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

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}