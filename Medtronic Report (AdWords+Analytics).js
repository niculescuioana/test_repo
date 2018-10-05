var PUSH_TRACKER_URL = 'https://docs.google.com/spreadsheets/d/1g21F1GGxNddqV-ivHhFqgAMuJKoDudPvztVuhqFsqh4/edit';
var PERFORMANCE_DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/1X8j5Q9FWCKkOQrPjgnsDy2SZI_vuAhqczUgUs9r3NYc/edit';


function main() {
  var CONFIG = parseInputs();
  var ids = Object.keys(CONFIG);
  
  MccApp.accounts().withIds(ids).executeInParallel('run', 'compile',JSON.stringify(CONFIG));
  
  /*var iter = MccApp.accounts().withIds(ids).get();
  while(iter.hasNext()) {
    MccApp.select(iter.next());
    run(JSON.stringify(CONFIG));
  }*/
}

function parseInputs() {
  var data = SpreadsheetApp.openByUrl(PUSH_TRACKER_URL).getSheetByName('Settings').getDataRange().getValues();
  data.shift();
  
  var CONFIG = {};
  var HEADER = ['NAME', 'ID', 'ANALYTICS_ID', 'GOAL_NUM', 'TAB_NAME'];
  for(var z in data) {
    if(!data[z][4]) { continue; }
    CONFIG[data[z][1]] = {};
    for(var k in HEADER) {
      CONFIG[data[z][1]][HEADER[k]] = data[z][k];
    }
  }
  
  return CONFIG;
}

function compile() {}

function run(input) {
  var CONFIG = JSON.parse(input)[AdWordsApp.currentAccount().getCustomerId()];
  compileAdWordsReport(CONFIG);
  
  //compileWeeklyAdWordsReport(CONFIG);
}

function compileAdWordsReport(CONFIG) {
  var TO = getAdWordsFormattedDate(0, 'yyyyMMdd');
  var stats = AdWordsApp.currentAccount().getStatsFor('20170701',TO);
  var ss = SpreadsheetApp.openByUrl(PERFORMANCE_DASHBOARD_URL);
  
  var idMap = {};
  var networkStats = {
    'Search': { 'Cost': 0, 'Clicks': 0, 'Impressions': 0, 'Goals': 0, 'Goals_2': 0, 'Transactions': 0, 'Revenue': 0 },
    'Display': { 'Cost': 0, 'Clicks': 0, 'Impressions': 0, 'Goals': 0, 'Goals_2': 0, 'Transactions': 0, 'Revenue': 0 }
  };
  
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = SEARCH').forDateRange('ALL_TIME').get();
  while(iter.hasNext()) {
    var camp = iter.next();
    idMap[camp.getId()] = 'Search'; 
    var campStats = camp.getStatsFor('20170701',TO);
    networkStats['Search']['Cost'] += campStats.getCost();
    networkStats['Search']['Clicks'] += campStats.getClicks();
    networkStats['Search']['Impressions'] += campStats.getImpressions();
    networkStats['Search']['Conversions'] += campStats.getConversions();
  }
  
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = DISPLAY').forDateRange('ALL_TIME').get();
  while(iter.hasNext()) {
    var camp = iter.next();
    idMap[camp.getId()] = 'Display'; 
    var campStats = camp.getStatsFor('20170701',TO);
    networkStats['Display']['Cost'] += campStats.getCost();
    networkStats['Display']['Clicks'] += campStats.getClicks();
    networkStats['Display']['Impressions'] += campStats.getImpressions();
    networkStats['Display']['Conversions'] += campStats.getConversions();    
  }
  
  var map = getStatsFromAnalytics('2017-07-01', getAdWordsFormattedDate(0, 'yyyy-MM-dd'), CONFIG, networkStats, idMap);
  
  var tab = ss.getSheetByName(CONFIG.TAB_NAME);
  tab.getRange(3,7,2,3).setValues([
    [stats.getCost(),networkStats['Search'].Cost,networkStats['Display'].Cost],
    [stats.getImpressions(),networkStats['Search'].Impressions,networkStats['Display'].Impressions]
  ]);
  
  tab.getRange(6,7,1,3).setValues([[stats.getClicks(),networkStats['Search'].Clicks,networkStats['Display'].Clicks]]);
  
  
  tab.getRange(14, 7, 3, 3).setValues([
    [map.Goals,networkStats['Search'].Goals,networkStats['Display'].Goals], 
    [map.Transactions,networkStats['Search'].Transactions,networkStats['Display'].Transactions], 
    [map.Revenue,networkStats['Search'].Revenue,networkStats['Display'].Revenue]
  ]);
  
  
  if(CONFIG.TAB_NAME == 'SA') {
    tab.getRange(17,7,1,3).setValues([[stats.getConversions(), networkStats['Search'].Conversions,networkStats['Display'].Conversions]]);
  }
  
  if(CONFIG.ANALYTICS_ID == 24930396 || CONFIG.ANALYTICS_ID == 111830734) {
    tab.getRange(15,7,1,3).setValues([[map['Goals_2'], networkStats['Search']['Goals_2'],networkStats['Display']['Goals_2']]]);
  }
  
  var accId = AdWordsApp.currentAccount().getCustomerId();
  var sheet = SpreadsheetApp.openByUrl(PUSH_TRACKER_URL).getSheetByName('AdWords');
  var data = sheet.getDataRange().getValues();
  data.shift();
  
  for(var z in data) {
    if(accId == data[z][1] ) {
      sheet.getRange(parseInt(z,10)+2, 5).setValue(stats.getCost());
    }
  }  
}

function getStatsFromAnalytics(FROM, TO, CONFIG, networkStats, idMap) {
  
  var goal = CONFIG.GOAL_NUM ?'ga:goal' + CONFIG.GOAL_NUM + 'Completions' :  'ga:goalCompletionsAll';
  
  if(CONFIG.ANALYTICS_ID == 24930396) {
    goal += ',ga:goal15Completions'; 
  } else if(CONFIG.ANALYTICS_ID == 111830734) {
    goal += ',ga:goal3Completions'; 
  }
  
  var results, attempts = 3;
  while(attempts > 0) {
    try {    
      results = Analytics.Data.Ga.get(
        'ga:'+CONFIG.ANALYTICS_ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue,"+goal,
        { 'dimensions': 'ga:adwordsCampaignID', 'filters': 'ga:source==google;ga:medium==cpc',
        'samplingLevel': 'HIGHER_PRECISION', 'max-results': 10000 });
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + CONFIG.ANALYTICS_ID);
      attempts--;
      Utilities.sleep(2000);
    }  
  }
  
  var stats = {
    'Transactions': 0, 'Revenue': 0, 'Goals': 0, 'Goals_2': 0
  };
  
  var rows = results.getRows();
  for(var k in rows) {
    if(rows[k][0] == '(not set)') { continue; }
    
    stats.Transactions += parseInt(rows[k][1],10);
    stats.Revenue += parseFloat(rows[k][2]);
    
    if(CONFIG.GOAL_NUM) {
      stats.Goals += parseInt(rows[k][3],10);
      if(rows[k][4]) {
        stats.Goals_2 += parseInt(rows[k][4],10);
      }
    }
    
    
    var network = '';
    if(idMap[rows[k][0]]) {
      network = idMap[rows[k][0]];
    }
    
    if(network) {
      networkStats[network].Transactions += parseInt(rows[k][1],10);
      networkStats[network].Revenue += parseFloat(rows[k][2]);
      
      if(CONFIG.GOAL_NUM) {
        networkStats[network].Goals += parseInt(rows[k][3],10);
        
        if(rows[k][4]) {
          networkStats[network].Goals_2 += parseInt(rows[k][4],10);
        }
      }
    }
    
  }
  
  return stats;
}


function compileWeeklyAdWordsReport(CONFIG) {
  var start = new Date('2017-07-03');
  start.setHours(12);
  
  var end = new Date('2018-03-11');
  start.setHours(12);
  
  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1lX-EmpZ9EU2mpa349ptE68cO12rtsYnBCPYjJoPk1Bc/edit#gid=1420766895')
  var tab = ss.getSheetByName(CONFIG.TAB_NAME);
  if(!tab){return; }
  var ROW_NUM = 2;
  while(start < end) {
    var FROM = Utilities.formatDate(start, 'PST', 'yyyy-MM-dd') ;
    start.setDate(start.getDate()+6);
    
    var TO = Utilities.formatDate(start, 'PST', 'yyyy-MM-dd') ;
    
    var stats = getWeeklyStatsFromAnalytics(FROM, TO, CONFIG);
    tab.getRange(ROW_NUM, 6).setValue(stats.Goals);
    start.setDate(start.getDate()+1);
    
    ROW_NUM++;
  }
}

function getWeeklyStatsFromAnalytics(FROM, TO, CONFIG) {
  
  var goal = CONFIG.GOAL_NUM ?'ga:goal' + CONFIG.GOAL_NUM + 'Completions' :  'ga:goalCompletionsAll';
  
  var results, attempts = 3;
  while(attempts > 0) {
    try {    
      results = Analytics.Data.Ga.get(
        'ga:'+CONFIG.ANALYTICS_ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue,"+goal,
        { 'dimensions': 'ga:adwordsCampaignID', 'filters': 'ga:source==google;ga:medium==cpc',
        'samplingLevel': 'HIGHER_PRECISION', 'max-results': 10000 });
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + CONFIG.ANALYTICS_ID);
      attempts--;
      Utilities.sleep(2000);
    }  
  }
  
  var stats = {
    'Transactions': 0, 'Revenue': 0, 'Goals': 0
  };
  
  var rows = results.getRows();
  for(var k in rows) {
    if(rows[k][0] == '(not set)') { continue; }
    
    stats.Transactions += parseInt(rows[k][1],10);
    stats.Revenue += parseFloat(rows[k][2]);
    
    if(CONFIG.GOAL_NUM) {
      stats.Goals += parseInt(rows[k][3],10);
    }
  }
  
  return stats;
}

/************************* Utils *******************************************************/

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}


function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}
