var url = 'https://docs.google.com/spreadsheets/d/1u82_Q7Xv--Xp2LszJIlPNGQU_PnkVlkoNyfSh0RDYA0/edit';
var PROFILE_ID = 335581;

function main() {
  MccApp.accounts().withCondition('Name = "Impact Factory"').executeInParallel('run');
}

function run() {
  runNormal();
  runCampaign('GENERIC CAMPAIGNS');
  runCampaign('OPEN/TAILORED CAMPAIGNS');
}

function runNormal() {
  var MONTH_YEAR = getAdWordsFormattedDate(0, 'MMMM yyyy');
  var MONTH = MONTH_YEAR.split(' ')[0];
  var YEAR = MONTH_YEAR.split(' ')[1];
  var END = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var START = END.substring(0,8) + '01';
  
  runForMonth(START, END, MONTH, YEAR);
  
  var dt = parseInt(END.split('-')[2], 10);
  if(dt == 1) {
    var MONTH_YEAR = getAdWordsFormattedDate(1, 'MMMM yyyy');
    var MONTH = MONTH_YEAR.split(' ')[0];
    var YEAR = MONTH_YEAR.split(' ')[1];
    var END = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
    var START = END.substring(0,8) + '01';
    
    runForMonth(START, END, MONTH, YEAR);
  }
}

function runForMonth(START, END, MONTH, YEAR) {
  var ss = SpreadsheetApp.openByUrl(url);
  var tab = ss.getSheetByName('Report - ' + YEAR);
  if(!tab) {
    var dummy = ss.getSheetByName('Report - ' + (YEAR-1));
    if(!dummy) { return; }
    
    ss.setActiveSheet(dummy);
    tab = ss.duplicateActiveSheet();
    tab.setName('Report - ' + YEAR);
    
    tab.getRange(4, 3, 4, tab.getLastColumn()-2).clearContent();
    tab.getRange(11, 3, 5, tab.getLastColumn()-2).clearContent();
    tab.getRange(20, 3, 1, tab.getLastColumn()-2).clearContent();
    tab.getRange(26, 3, 4, tab.getLastColumn()-2).clearContent();
  }
  
  var stats = AdWordsApp.currentAccount().getStatsFor(START.replace(/-/g, ''), END.replace(/-/g, ''));
  
  var data = tab.getDataRange().getValues();
  data.shift();
  var header = data.shift();
  var col = header.indexOf(MONTH) + 1;
  tab.getRange(26, col, 4, 1).setValues([
    [stats.getClicks()],
    [stats.getImpressions()],
    [stats.getCtr()],
    [stats.getAverageCpc()]
  ]);
  
  tab.getRange(20, col).setValue(stats.getCost());
  
  var optArgs = { 
    'filters': 'ga:source==google;ga:medium==cpc',
    //'segment': 'users::condition::ga:source==google;users::condition::ga:medium=@cpc'
  };
  
  var stats = getDataFromAnalytics(optArgs, START, END);
  tab.getRange(4, col, 4, 1).setValues([
    [stats['Goal10']], [stats['Goal11']],
    [stats['AssistedGoal10']], [stats['AssistedGoal11']]
  ]);
  
  tab.getRange(11, col, 5, 1).setValues([
    [stats['Transactions']], [stats['AssistedTransactions']], 
    [stats['TransactionRevenue']], [stats['AssistedTransactionRevenue']],
    [stats['TransactionRevenue'] + stats['AssistedTransactionRevenue']]
  ]);
}


function getDataFromAnalytics(optArgs, FROM, TO) {
  var stats = {
    'Transactions': 0, 'TransactionRevenue': 0,  'Goal10': 0,  'Goal11': 0,
    'AssistedTransactions': 0, 'AssistedTransactionRevenue': 0,  'AssistedGoal10': 0,  'AssistedGoal11': 0
  };
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue,ga:goal10Completions,ga:goal11Completions",
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
    stats['Transactions'] += parseInt(rows[k][0],10);
    stats['TransactionRevenue'] += parseFloat(rows[k][1]);
    stats['Goal10'] += parseInt(rows[k][2],10);
    stats['Goal11'] += parseInt(rows[k][3],10);
  }
  
  
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search'];
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType,mcf:conversionGoalNumber',
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                  // End-date (format yyyy-MM-dd).
    "mcf:totalConversions,mcf:totalConversionValue",
    optArgs
  );
  
  var rows = results.rows;
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    if(rows[k][1].primitiveValue == 'Transaction') {
      stats.AssistedTransactions += parseInt(rows[k][3].primitiveValue,10);
      stats.AssistedTransactionRevenue += parseFloat(rows[k][4].primitiveValue);      
    } else if(rows[k][2].primitiveValue == '010') {
      stats['AssistedGoal10'] += parseInt(rows[k][3].primitiveValue,10);
    } else if(rows[k][2].primitiveValue == '011') {
      stats['AssistedGoal11'] += parseInt(rows[k][3].primitiveValue,10);
    }
  }
  
  
  return stats;
}

function runCampaign(CAMPAIGN_LABEL) {
  var MONTH_YEAR = getAdWordsFormattedDate(0, 'MMMM yyyy');
  var MONTH = MONTH_YEAR.split(' ')[0];
  var YEAR = MONTH_YEAR.split(' ')[1];
  var END = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var START = END.substring(0,8) + '01';
  
  runForMonthCampaign(START, END, MONTH, YEAR, CAMPAIGN_LABEL);
  
  var dt = parseInt(END.split('-')[2], 10);
  if(dt == 1) {
    var MONTH_YEAR = getAdWordsFormattedDate(1, 'MMMM yyyy');
    var MONTH = MONTH_YEAR.split(' ')[0];
    var YEAR = MONTH_YEAR.split(' ')[1];
    var END = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
    var START = END.substring(0,8) + '01';
    
    runForMonthCampaign(START, END, MONTH, YEAR, CAMPAIGN_LABEL);
  }
}

function runForMonthCampaign(START, END, MONTH, YEAR, CAMPAIGN_LABEL) {
  var ss = SpreadsheetApp.openByUrl(url);
  var tab = ss.getSheetByName(CAMPAIGN_LABEL + ' - ' + YEAR);
  if(!tab) {
    var dummy = ss.getSheetByName(CAMPAIGN_LABEL + ' - ' + (YEAR-1));
    if(!dummy) { return; }
    
    ss.setActiveSheet(dummy);
    tab = ss.duplicateActiveSheet();
    tab.setName(CAMPAIGN_LABEL + ' - ' + YEAR);
    
    tab.getRange(4, 3, 4, tab.getLastColumn()-2).clearContent();
    tab.getRange(11, 3, 5, tab.getLastColumn()-2).clearContent();
    tab.getRange(20, 3, 1, tab.getLastColumn()-2).clearContent();
    tab.getRange(26, 3, 4, tab.getLastColumn()-2).clearContent();
  }
  
  var campMap = {}, clicks = 0, impressions = 0, cost = 0;
  var iter = AdWordsApp.campaigns().withCondition('LabelNames CONTAINS_ANY ["' + CAMPAIGN_LABEL + '"]').get();
  while(iter.hasNext()) {
    var camp = iter.next();
    var stats = camp.getStatsFor(START.replace(/-/g, ''), END.replace(/-/g, ''));
    clicks += stats.getClicks();
    impressions += stats.getImpressions();
    cost += stats.getCost();
    campMap[camp.getName()] = 1;
  }
  
  var iter = AdWordsApp.videoCampaigns().withCondition('LabelNames CONTAINS_ANY ["' + CAMPAIGN_LABEL + '"]').get();
  while(iter.hasNext()) {
    var camp = iter.next();
    var stats = camp.getStatsFor(START.replace(/-/g, ''), END.replace(/-/g, ''));
    clicks += stats.getClicks();
    impressions += stats.getImpressions();
    cost += stats.getCost();
    campMap[camp.getName()] = 1;
  }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  var header = data.shift();
  var col = header.indexOf(MONTH) + 1;
  tab.getRange(26, col, 4, 1).setValues([
    [clicks],
    [impressions],
    [impressions == 0 ? 0 : clicks / impressions],
    [clicks == 0 ? 0 : cost / clicks]
  ]);
  
  tab.getRange(20, col).setValue(cost);
  
  var optArgs = { 
    'dimensions': 'ga:campaign',
    'filters': 'ga:source==google;ga:medium==cpc'
  };
  
  var stats = getDataFromAnalyticsCampaign(optArgs, START, END, campMap);
  tab.getRange(4, col, 4, 1).setValues([
    [stats['Goal10']], [stats['Goal11']],
    [stats['AssistedGoal10']], [stats['AssistedGoal11']]
  ]);
  
  tab.getRange(11, col, 5, 1).setValues([
    [stats['Transactions']], [stats['AssistedTransactions']], 
    [stats['TransactionRevenue']], [stats['AssistedTransactionRevenue']],
    [stats['TransactionRevenue'] + stats['AssistedTransactionRevenue']]
  ]);
}


function getDataFromAnalyticsCampaign(optArgs, FROM, TO, campaignMap) {
  var stats = {
    'Transactions': 0, 'TransactionRevenue': 0,  'Goal10': 0,  'Goal11': 0,
    'AssistedTransactions': 0, 'AssistedTransactionRevenue': 0,  'AssistedGoal10': 0,  'AssistedGoal11': 0
  };
  
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:transactions,ga:transactionRevenue,ga:goal10Completions,ga:goal11Completions",
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
    var camp = rows[k][0];
    if(!campaignMap[camp]) { continue; }
    stats['Transactions'] += parseInt(rows[k][1],10);
    stats['TransactionRevenue'] += parseFloat(rows[k][2]);
    stats['Goal10'] += parseInt(rows[k][3],10);
    stats['Goal11'] += parseInt(rows[k][4],10);
  }
  
  
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search'];
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType,mcf:conversionGoalNumber,mcf:adwordsCampaignPath',
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                  // End-date (format yyyy-MM-dd).
    "mcf:totalConversions,mcf:totalConversionValue",
    optArgs
  );
  
  var rows = results.rows;
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    var camp = JSON.parse(rows[k][3]).conversionPathValue[0].nodeValue;
    if(!campaignMap[camp]) { continue; }
    
    if(rows[k][1].primitiveValue == 'Transaction') {
      stats.AssistedTransactions += parseInt(rows[k][4].primitiveValue,10);
      stats.AssistedTransactionRevenue += parseFloat(rows[k][5].primitiveValue);      
    } else if(rows[k][2].primitiveValue == '010') {
      stats['AssistedGoal10'] += parseInt(rows[k][4].primitiveValue,10);
    } else if(rows[k][2].primitiveValue == '011') {
      stats['AssistedGoal11'] += parseInt(rows[k][4].primitiveValue,10);
    }
  }
  
  
  return stats;
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}
