var REPORT_URL = 'https://docs.google.com/spreadsheets/d/12sCEEb8rR3E-QLwATLVe7RXzPb-wTobTISl4dnkYIlw/edit';

var CONFIG = {
  '197-924-6181': { 'PROFILE_ID': 119695269, 'GOALS': 'ga:goal1Completions,ga:goal3Completions,ga:goal5Completions' },
  //'107-044-7779': { 'PROFILE_ID': 12331684 }
}

function main() {
  MccApp.accounts().withIds(Object.keys(CONFIG)).executeInParallel('runScript');
}

function runScript() {
  var SETTINGS = CONFIG[AdWordsApp.currentAccount().getCustomerId()];
  var accName = AdWordsApp.currentAccount().getName();
  if(accName == 'Aspect') {
    accName = 'Aspect UK';
  }    
  
  var initMap = {
    'Search': { 'Clicks': 0, 'Impressions': 0, 'Ctr': 0, 'Cost': 0, 'Conversions': 0  },
    'Display': { 'Clicks': 0, 'Impressions': 0, 'Ctr': 0, 'Cost': 0, 'Conversions': 0  },
    'YouTube': { 'Clicks': 0, 'Impressions': 0, 'Ctr': 0, 'Cost': 0, 'Conversions': 0  },
  };
  
  var today = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
  today.setHours(12);
  
  var end = Utilities.formatDate(today, 'PST', 'yyyyMMdd');
  SETTINGS.TO = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');

  today.setDate(1);
  today.setDate(0);
  today.setDate(1);
  var start = Utilities.formatDate(today, 'PST', 'yyyyMMdd');
  SETTINGS.FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
  
  SETTINGS.DATE_RANGE = start + ',' + end;
  
  var OPTIONS = { includeZeroImpressions : true };
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var cols = ['MonthOfYear','Year','AdNetworkType1','Impressions','Clicks',
              'Cost','Ctr','Conversions'];
  
  var query = ['select',cols.join(','),'from',report,
               'during', SETTINGS.DATE_RANGE].join(' ');
  
  var stats = {};
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    row.AdNetworkType1 = row.AdNetworkType1.split(' ')[0];
    if(!initMap[row.AdNetworkType1]) { continue; }
    
    if(!stats[row.Year]) {
      stats[row.Year] = {};
    }
    
    var month = row.MonthOfYear.substring(0,3);
    if(!stats[row.Year][month]) {
      stats[row.Year][month] = JSON.parse(JSON.stringify(initMap));
    }      
    
    stats[row.Year][month][row.AdNetworkType1].Clicks = parseInt(row.Clicks,10);
    stats[row.Year][month][row.AdNetworkType1].Impressions = parseInt(row.Impressions,10);
    stats[row.Year][month][row.AdNetworkType1].Cost = parseFloat(row.Cost.toString().replace(/,/g,''));    
    stats[row.Year][month][row.AdNetworkType1].Conversions = parseFloat(row.Conversions.toString().replace(/,/g,''));        
    stats[row.Year][month][row.AdNetworkType1].Ctr = row.Ctr;
  }
  
  var ss = SpreadsheetApp.openByUrl(REPORT_URL);
  for(var year in stats) {
    var tabName = [accName,year,'MQL'].join(' ');
    var tab = ss.getSheetByName(tabName);
    if(!tab) { Logger.log(tabName); continue; }
    for(var month in stats[year]) {
      var row = stats[year][month];
      var data = tab.getDataRange().getValues();
      data.shift();
      var header = data.shift();
      var col = header.indexOf(month)+1;
      if(col === 0) { continue; }
      
      //Logger.log(row['Search Network'].Cost);
      tab.getRange(3,col,3,1).setValues([[row['Search'].Cost],[row['Display'].Cost],[row['YouTube'].Cost]]);
      
      tab.getRange(23,col,7,1).setValues([[row['Search'].Clicks],[row['Search'].Ctr], [row['Search'].Impressions],
                                         [''], [row['Display'].Clicks],[row['Display'].Ctr], [row['Display'].Impressions]]);
    }
  }
  
  var optArgs = { 
    'dimensions': 'ga:year,ga:month', 
    'samplingLevel': 'HIGHER_PRECISION', 
    'filters': 'ga:source==google;ga:medium==cpc' 
  };
  
  var metrics = SETTINGS.GOALS;
  var goalCompletionsMap = getAnalyticsStats(SETTINGS, optArgs, metrics);
  
  for(var year in goalCompletionsMap) {
    var tabName = [accName,year,'MQL'].join(' ');
    var tab = ss.getSheetByName(tabName);
    if(!tab) { Logger.log(tabName); continue; }
    for(var month in goalCompletionsMap[year]) {
      var data = tab.getDataRange().getValues();
      data.shift();
      var header = data.shift();
      var col = header.indexOf(month)+1;
      if(col === 0) { continue; }
      
      tab.getRange(34,col).setValue(goalCompletionsMap[year][month]['A']);
      tab.getRange(35,col).setValue(goalCompletionsMap[year][month]['B']);      
      tab.getRange(37,col).setValue(goalCompletionsMap[year][month]['C']);            
    }
  }
  
  var assistedConversionMap = getDataFromMCF(SETTINGS);
  for(var year in assistedConversionMap) {
    var tabName = [accName,year,'MQL'].join(' ');
    var tab = ss.getSheetByName(tabName);
    if(!tab) { Logger.log(tabName); continue; }
    for(var month in assistedConversionMap[year]) {
      var data = tab.getDataRange().getValues();
      data.shift();
      var header = data.shift();
      var col = header.indexOf(month)+1;
      if(col === 0) { continue; }
      
      tab.getRange(36,col).setValue(assistedConversionMap[year][month]);
    }
  }
}

function getAnalyticsStats(SETTINGS, optArgs, metrics) {
  //log(JSON.stringify(optArgs));
  var attempts = 3;
  var results;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      results = Analytics.Data.Ga.get(
        'ga:'+SETTINGS.PROFILE_ID,                   // Table id (format ga:xxxxxx).
        SETTINGS.FROM,                 // Start-date (format yyyy-MM-dd).
        SETTINGS.TO,                   // End-date (format yyyy-MM-dd).
        metrics,
        optArgs);
      break;
    } catch(ex) {
      log(ex + " ID: " + SETTINGS.PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul',
                    'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
  var stats = {};
  var rows = results.getRows();
  //log(rows);
  for(var k in rows) {
    if(!stats[rows[k][0]]) {
      stats[rows[k][0]] = {};
    }
    
    var month = monthNames[parseInt(rows[k][1],10)];
    if(!stats[rows[k][0]][month]) {
      stats[rows[k][0]][month] = {'A': 0, 'B': 0, 'C': 0};
    }
    
    stats[rows[k][0]][month]['A'] += parseInt(rows[k][2],10);
    stats[rows[k][0]][month]['B'] += parseInt(rows[k][3],10);
    stats[rows[k][0]][month]['C'] += parseInt(rows[k][4],10); 
  }
  
  return stats;
}


function getDataFromMCF(SETTINGS) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search'];
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionDate', "samplingLevel": "HIGHER_PRECISION",
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+SETTINGS.PROFILE_ID,                   // Table id (format ga:xxxxxx).
    SETTINGS.FROM,                 // Start-date (format yyyy-MM-dd).
    SETTINGS.TO,                   // End-date (format yyyy-MM-dd).
    "mcf:assistedConversions",
    optArgs
  );
  var monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul',
                    'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  
  var stats = {};
  var rows = results.getRows();
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    var year = rows[k][1]["primitiveValue"].substring(0,4);
    var month = monthNames[parseInt(rows[k][1]["primitiveValue"].substring(4,6),10)];
    
    if(!stats[year]) {
      stats[year] = {};
    }
    
    if(!stats[year][month]) {
      stats[year][month] = 0;
    }
    
    stats[year][month] += parseInt(rows[k][2]["primitiveValue"],10);
  }
  
  return stats;
}

function log(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg); 
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}