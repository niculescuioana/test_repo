/******************************************
* Report - Daily Benchmark Report
* @version: 5.0
* @author: Naman Jindal (nj.itprof@gmail.com)
******************************************/

var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1969885541';
var TAB_NAME = 'Benchmark Urls';

function main() {
  
  //var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  //data.shift();
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var hour = NOW.getHours();
  
  //var index = hour%4;
  
  var SETTINGS_MAP = readInputsForAccounts();
  //Logger.log('Running for: ' + data[index][0]);
  
  var URL = 'https://docs.google.com/spreadsheets/d/1rkS8Tb4Wscior-KKqn4fcWygLVMn1e9U2OiljzzQNJc/edit';
  cleanSheets('Harsh', URL);

  //var LABELS = data[index][0].split(',');
  
  var accountIter = MccApp.accounts()
  .withCondition('LabelNames CONTAINS "Harsh"')
  .orderBy('Impressions DESC')
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
    
    SETTINGS.LABEL = 'Harsh';
    SETTINGS.REPORT_URL = URL;
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
  
  if(sheet.getLastRow() > 136) {
    compileEcommerceStats(SETTINGS,sheet,stats,dateKeys); 
  }
  
  log('Finished');
}

function compileEcommerceStats(SETTINGS,sheet,stats,dateKeys) {
  log('Compiling Ecommerce Report');
  
  sheet.getRange(139,2,stats.rtos.length,2).setValues(stats.rtos).sort({column: 2, ascending: true});
  
  var PROFILE_IDS = SETTINGS.PROFILE_IDS;
  if(!PROFILE_IDS) { log('No PROFILE_IDS'); return; }
  
  var analyticsStats = { revenue: [], adRevenue: [] }
  
  readAnalyticsStats(analyticsStats,PROFILE_IDS);
  
  if(analyticsStats.revenue.length > 0) {
    sheet.getRange(156,2,analyticsStats.revenue.length,2).setValues(analyticsStats.revenue).sort({column: 2, ascending: true});
  }
  
  if(analyticsStats.adRevenue.length > 0) {
    sheet.getRange(173,2,analyticsStats.adRevenue.length,2).setValues(analyticsStats.adRevenue).sort({column: 2, ascending: true});  
  }
  
  var shoppingStatsMap = getShoppingStats();
  stats.shoppingClicks = [];
  for(var key in shoppingStatsMap) {
    stats.shoppingClicks.push([key, shoppingStatsMap[key].Clicks, shoppingStatsMap[key].AverageClicks]);
  }
  
  if(stats.shoppingClicks.length == 0) {
    for(var key in dateKeys) {
      stats.shoppingClicks.push([key, 0, 0]);
    }
  }
  
  sheet.getRange(190,2,stats.shoppingClicks.length,3).setValues(stats.shoppingClicks).sort({column: 2, ascending: true});    
}

function getShoppingStats(){ 
  var shoppingStatsMap = {};
  var shoppingCampaignIds = [];
  
  var avgClicks = getShoppingCampaignIds(shoppingCampaignIds);
  if(shoppingCampaignIds.length == 0) { return shoppingStatsMap; }
  
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['Date','CampaignId','Impressions','Clicks','Conversions','Cost',
              'AveragePosition','CostPerConversion','ConversionValue'];
  var report = 'CAMPAIGN_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where CampaignId IN [' + shoppingCampaignIds.join(',') + ']',
               'during','LAST_14_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(!shoppingStatsMap[row.Date]) {
      shoppingStatsMap[row.Date] = { 'Clicks': 0, 'AverageClicks': avgClicks }
    }
    
    shoppingStatsMap[row.Date].Clicks += parseInt(row.Clicks,10);
  } 
  
  return shoppingStatsMap;
}

function getShoppingCampaignIds(shoppingCampaignIds) {
  var totalShoppingClicks = 0;
  var iter = AdWordsApp.shoppingCampaigns()
  .withCondition('Clicks > 0')
  .forDateRange('LAST_30_DAYS')
  .get();
  
  while(iter.hasNext()) {
    var camp = iter.next();
    shoppingCampaignIds.push(camp.getId());
    totalShoppingClicks += camp.getStatsFor('LAST_30_DAYS').getClicks();
  }
  
  var avgClicks = totalShoppingClicks/30;
  return avgClicks;
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



/************* Analytics Code ************************/

function readAnalyticsStats(analyticsStats,PROFILE_IDS) {
  var ids =  PROFILE_IDS.toString().split(',');
  
  if(ids.length == 0) { log('No profiles found'); }
  
  var FROM = getAdWordsFormattedDate(14, 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var stats = {};
  
  var optArgs = { 'dimensions': 'ga:date,ga:channelGrouping' }
  var optArgs_1 = { 'dimensions': 'ga:date,ga:channelGrouping', 'filters': 'ga:medium==cpc' };
  var results = {};
  for(var k in ids) {
    getAnalyticsStatsByDate(ids[k].trim(),stats,FROM,TO,optArgs,optArgs_1);
  }
  
  for(var key in stats) {
    var date = key.substring(0,4) + '/' + key.substring(4,6) + '/' + key.substring(6,8);
    analyticsStats.revenue.push([date, stats[key].Revenue]);
    analyticsStats.adRevenue.push([date, stats[key].AdRevenue]);    
  }
}


//Queries Analyltics stats from API
function getAnalyticsStatsByDate(id, stats, FROM, TO, optArgs, optArgs_1 ) {
  var tryAgain = false;
  try {
    fetchAnalyticsStats(id, stats, FROM, TO, optArgs, optArgs_1);
  } catch(e) {
    var exc = (id+': '+e);
    log(exc);
    if(e.message.toLowerCase().indexOf('user limit') > -1 ||
       e.message.toLowerCase().indexOf('user rate') > -1) { 
      tryAgain = true; 
    }
  }
  
  if(tryAgain) {
    Utilities.sleep(1100);
    log('Retrying');    
    try {
      fetchAnalyticsStats(id, stats, FROM, TO, optArgs, optArgs_1);
    } catch(e) {
      var exc = (id+': '+e);
      log(exc);
    } 
  }
  
}

function fetchAnalyticsStats(id, stats, FROM, TO, optArgs, optArgs_1) {
  
  // Make a requst to the API.
  var results = Analytics.Data.Ga.get(
    'ga:'+id,              // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                   // End-date (format yyyy-MM-dd).
    "ga:transactionRevenue",
    optArgs);
  
  var rows = results.getRows();
  for(var k in rows) {
    if(!stats[rows[k][0]]) { stats[rows[k][0]] = { Revenue: 0, AdRevenue: 0 } }
    stats[rows[k][0]].Revenue += parseFloat(rows[k][2]);
    
    if(rows[k][1] == 'Paid Search') {
      stats[rows[k][0]].AdRevenue += parseFloat(rows[k][2]);
    }
  }
  
  /**
  // Make a request to the API.
  var results = Analytics.Data.Ga.get(
  'ga:'+id,              // Table id (format ga:xxxxxx).
  FROM,                 // Start-date (format yyyy-MM-dd).
  TO,                   // End-date (format yyyy-MM-dd).
  "ga:transactionRevenue",
  optArgs_1);
  
  var rows = results.getRows();
  for(var k in rows) {
  if(!stats[rows[k][0]]) { stats[rows[k][0]] = { Revenue: 0, AdRevenue: 0 } }
  stats[rows[k][0]].AdRevenue += parseFloat(rows[k][2]);
  }
  **/
}

function getLinkedProfileIds(UA_NUMBER) {  
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName('Analytics - All Profiles').getDataRange().getValues();
  data.shift();
  
  var profiles = [];
  for(var k in data) {
    if(UA_NUMBER.indexOf(data[k][2]) > -1) {
      profiles.push(data[k][3]);
    }
  }
  
  return profiles;
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}