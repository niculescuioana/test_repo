var URL = 'https://docs.google.com/spreadsheets/d/1WAL2jXKQtK7OuEUZU0yq8XVUZM_1J6_PHFE3u3yGSug/edit#gid=1074707401';

function main() {
  MccApp.accounts().withIds(['486-718-6575']).executeInParallel('run')
}

function run() {
  compileOverallReport();
  compileMicrositesReport();
}

function compileMicrositesReport() {
  var PROFILE_MAP = {
    'Breaker Finder': 125432776,
    'Circuit Analyzer': 125440437, 
    'Clamp Meter': 125444523, 
    'Circuit Tracer': 125380488
  };
  
  
  var year = getAdWordsFormattedDate(0, 'yyyy');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var FROM = TO.split('-')[0] + '-01-01';
  
  var analyticsStats = { 'all': {} };
  
  for(var LABEL in PROFILE_MAP) {
    getDataFromMicrositeAnalytics(FROM, TO, PROFILE_MAP[LABEL], analyticsStats, LABEL);
  }
  
  var monthNames = ['Januray', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
  var ss = SpreadsheetApp.openByUrl(URL);
  
  var LABELS = Object.keys(PROFILE_MAP)
  LABELS.push('all');
  var sheet = ss.getSheetByName('Microsites ' + year);
  var col = 1;
  for(var z in LABELS) {
    var out1 = [], out2 = [], out3 = [];
    var stats = analyticsStats[LABELS[z]];
    for(var i in monthNames) {
      var month = monthNames[i];
      if(!stats[month]) {
        stats[month] = { 'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0 } 
      }
      
      out1.push([month, stats[month].Sessions]);
      
      var avgSD = stats[month].Sessions == 0 ? 0 : round(stats[month].SessionDuration / stats[month].Sessions,2);
      out2.push([month, avgSD]);
      
      var br = stats[month].Sessions == 0 ? 0 : round(100*stats[month].Bounces / stats[month].Sessions,2)+'%';
      out3.push([month, br]);
    }
    
    sheet.getRange(3, col, out1.length, out1[0].length).setValues(out1);
    sheet.getRange(17, col, out2.length, out2[0].length).setValues(out2);
    sheet.getRange(31, col, out3.length, out3[0].length).setValues(out3);
    
    col += 3;
  }
}


function getDataFromMicrositeAnalytics(FROM, TO, ID, map, label) {
  var monthNames = ['', 'Januray', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
  
  
  var results, attempts = 3;
  while(attempts > 0) {
    try {    
      results = Analytics.Data.Ga.get(
        'ga:'+ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:sessions,ga:users,ga:goal1Completions,ga:goal2Completions,ga:goal10Completions,ga:goal11Completions,ga:goal12Completions,ga:goal13Completions,ga:sessionDuration,ga:bounces",
        { 'dimensions': 'ga:month,ga:adwordsCampaignID', 
        'samplingLevel': 'HIGHER_PRECISION', 'sort': 'ga:month' });
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + ID);
      attempts--;
      Utilities.sleep(2000);
    }  
  }
 
  var rows = results.getRows();
  for(var k in rows) {
    var month = monthNames[parseInt(rows[k][0],10)];
    
    if(!map['all'][month]) {
      map['all'][month] = {
        'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0
      };
    }
    
    var conversions = parseInt(rows[k][4],10) + parseInt(rows[k][5],10) + parseInt(rows[k][6],10) + parseInt(rows[k][7],10) + parseInt(rows[k][8],10) + parseInt(rows[k][9],10);
    map['all'][month].Sessions += parseInt(rows[k][2],10);
    map['all'][month].Users += parseInt(rows[k][3],10);
    map['all'][month].Conversions += conversions;
    map['all'][month].SessionDuration += parseFloat(rows[k][10]);
    map['all'][month].Bounces += parseInt(rows[k][11],10);
    
    
    if(!map[label]) {
      map[label] = {}; 
    }
    
    if(!map[label][month]) {
      map[label][month] = {
        'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0
      }
    }
    
    map[label][month].Sessions += parseInt(rows[k][2],10);
    map[label][month].Users += parseInt(rows[k][3],10);
    map[label][month].Conversions += conversions;
    map[label][month].SessionDuration += parseFloat(rows[k][10]);
    map[label][month].Bounces += parseInt(rows[k][11],10);
  }
  
  return map;
}



function compileOverallReport() {
  var LABELS = ['Product - Breaker Finder', 'Product - Circuit Analyzer' , 'Product - Clamp Meter', 'Product - Circuit Tracer'];
  
  var campMap = {};
  for(var z in LABELS) {
    var iter = AdWordsApp.campaigns().withCondition('LabelNames CONTAINS_ANY ["' + LABELS[z] + '"]').get();
    while(iter.hasNext()) {
      var camp = iter.next();
      campMap[camp.getId()] = LABELS[z];
    }
  }
  
  var year = getAdWordsFormattedDate(0, 'yyyy');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var FROM = TO.split('-')[0] + '-01-01';
  var analyticsStats = getDataFromAnalytics(FROM, TO, campMap);
  //Logger.log(JSON.stringify(analyticsStats));
  var allStats = analyticsStats['all'];
  
  var out = [];
  for(var month in allStats) {
    var row = allStats[month];
    out.push([month, row.Users, '', month, row.Conversions, '', month, row.Sessions]);
  }
  
  var monthNames = ['Januray', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
  var ss = SpreadsheetApp.openByUrl(URL);
  var tab = ss.getSheetByName(year);
  tab.getRange(4, 1, out.length, out[0].length).setValues(out);
  
  
  LABELS.push('all');
  var sheet = ss.getSheetByName('Landing Pages ' + year);
  var col = 1;
  for(var z in LABELS) {
    var out1 = [], out2 = [], out3 = [];
    var stats = analyticsStats[LABELS[z]];
    for(var i in monthNames) {
      var month = monthNames[i];
      if(!stats[month]) {
        stats[month] = { 'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0 } 
      }
      
      out1.push([month, stats[month].Sessions]);
      
      var avgSD = stats[month].Sessions == 0 ? 0 : round(stats[month].SessionDuration / stats[month].Sessions,2);
      out2.push([month, avgSD]);
      
      var br = stats[month].Sessions == 0 ? 0 : round(100*stats[month].Bounces / stats[month].Sessions,2)+'%';
      out3.push([month, br]);
    }
    
    sheet.getRange(3, col, out1.length, out1[0].length).setValues(out1);
    sheet.getRange(17, col, out2.length, out2[0].length).setValues(out2);
    sheet.getRange(31, col, out3.length, out3[0].length).setValues(out3);
    
    col += 3;
  }
}

function getDataFromAnalytics(FROM, TO, campMap) {
  var monthNames = ['', 'Januray', 'February', 'March', 'April', 'May', 'June',
                    'July',' August', 'September', 'October', 'November', 'December'];
  
  var ID = 101539011, results, attempts = 3;
  while(attempts > 0) {
    try {    
      results = Analytics.Data.Ga.get(
        'ga:'+ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:sessions,ga:users,ga:goal1Completions,ga:goal2Completions,ga:goal10Completions,ga:goal11Completions,ga:goal12Completions,ga:goal13Completions,ga:sessionDuration,ga:bounces",
        { 'dimensions': 'ga:month,ga:adwordsCampaignID', 
        'samplingLevel': 'HIGHER_PRECISION', 'sort': 'ga:month' });
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + ID);
      attempts--;
      Utilities.sleep(2000);
    }  
  }
  
  var map = { 'all': {} };
  var rows = results.getRows();
  for(var k in rows) {
    var month = monthNames[parseInt(rows[k][0],10)];
    
    if(!map['all'][month]) {
      map['all'][month] = {
        'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0
      };
    }
    
    var conversions = parseInt(rows[k][4],10) + parseInt(rows[k][5],10) + parseInt(rows[k][6],10) + parseInt(rows[k][7],10) + parseInt(rows[k][8],10) + parseInt(rows[k][9],10);
    map['all'][month].Sessions += parseInt(rows[k][2],10);
    map['all'][month].Users += parseInt(rows[k][3],10);
    map['all'][month].Conversions += conversions;
    map['all'][month].SessionDuration += parseFloat(rows[k][10]);
    map['all'][month].Bounces += parseInt(rows[k][11],10);
    
    
    var label = campMap[rows[k][1]];
    if(!label) { continue; }
    
    if(!map[label]) {
      map[label] = {}; 
    }
    
    if(!map[label][month]) {
      map[label][month] = {
        'Sessions': 0,  'Users': 0, 'Conversions': 0, 'SessionDuration': 0, 'Bounces': 0
      }
    }
    
    map[label][month].Sessions += parseInt(rows[k][2],10);
    map[label][month].Users += parseInt(rows[k][3],10);
    map[label][month].Conversions += conversions;
    map[label][month].SessionDuration += parseFloat(rows[k][10]);
    map[label][month].Bounces += parseInt(rows[k][11],10);
  }
  
  return map;
}


function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getName(),format);
}


function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}