/*************************************************
* E-Commerce Report - Analytics
* @version: 1.1
* @author: Naman Jindal (naman@pushgroup.co.uk)
***************************************************/

var MASTER_DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1415541725';
var DASHBOARD_URLS_TAB = 'Dashboard Urls';
var ANALYTICS_TAB_NAME = 'Analytics - All Profiles';

var INPUT_TAB_NAME = 'Management Report';
var TEMPLATE_URL = 'https://docs.google.com/spreadsheets/d/1dUQfRXjjCNMEm_J74CMvF6qft7T2TYqPgMjQUAqWYq4/edit';

var FOLDER_ID = '0ByGcUDAdVlY5WEpEajZ4bWJKUEU';

function main() {
  var masterSS = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL);
  var data = [
    ['Neeraj', 'https://docs.google.com/spreadsheets/d/1WgnrCuWl6cX_8wlWAHCIF_qH1CSr1AsUcHeY5TE7p64/edit'],
    ['Ian', 'https://docs.google.com/spreadsheets/d/1TJNezTn0QIoWBCOt_iRCaxO2Yw9v9Zaffhste1klTQQ/edit'],
    ['Jay', 'https://docs.google.com/spreadsheets/d/1HSdATgm_IHStt0lwpvlxhKK1GYkv1_iq1Ys5fTxnB8I/edit'],
    ['Mike', 'https://docs.google.com/spreadsheets/d/1FK6RELdZd7tFmDFb8KvDEzUqg7jvlLqcB0OCR2Sf_TM/edit']
  ]
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var hour = NOW.getHours();
  
  var index = hour%4;
  //index=5;
  var LABEL = data[index][0];
  var REPORT_URL = data[index][1];
  Logger.log('Running for: ' + LABEL);
  //return;
  //if(LABEL != 'Neeraj' && LABEL != 'Jay' && LABEL != 'Ian' && LABEL != 'Mike') { return; }
  
  var SETTINGS = parseInputs(REPORT_URL);
  var names = Object.keys(SETTINGS);

  var iter = MccApp.accounts()
  .withCondition('LabelNames CONTAINS "'+LABEL+'"')
  .withCondition('Name IN ["' + names.join('","') + '"]')
  .get();
  
  while(iter.hasNext()) {
   var account = iter.next();
    MccApp.select(account);
    runScript(JSON.stringify(SETTINGS));
  }
  
  //.withCondition('Name = "Love Lula"')
  //.executeInParallel('runScript','compileResults',JSON.stringify(SETTINGS));
}

function compileResults() {
  // Do something here 
  Logger.log('Finished');
}

function parseInputs(REPORT_URL) {
  var data = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(INPUT_TAB_NAME).getDataRange().getValues();
  var header = data.shift();
  data.shift();
  
  var uaIndex = header.indexOf('ANALYTICS REPORTING');
  var urlIndex = header.indexOf('ECOMMERCE REPORT');
  
  var SETTINGS = {};
  for(var k in data) {
    //if(!data[k][urlIndex]) { continue; }
    SETTINGS[data[k][0]] = { 
      'URL': data[k][urlIndex], 
      'ROW_NUM': parseInt(k,10)+3, 'COL_NUM': parseInt(urlIndex,10)+2,
      'DASHBOARD_URL': REPORT_URL 
    };
  }
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Account Inputs');
  var inputData = sheet.getDataRange().getValues();
  var scriptNameHeader = inputData.shift();
  var inputHeader = inputData.shift();
  
  for(var j in inputData){
    if(!SETTINGS[inputData[j][0]]) {
      continue;
    }
    
    for(var l in inputHeader) {
      SETTINGS[inputData[j][0]][inputHeader[l]] = inputData[j][l];
    }
  }
  
  return SETTINGS;
}

function runScript(INPUT) {
  var SETTINGS = JSON.parse(INPUT)[AdWordsApp.currentAccount().getName()];
  if(!SETTINGS) { return; }
  
  var ss, createSpreadsheet = false;
  try {
    ss = SpreadsheetApp.openByUrl(SETTINGS.URL);
  } catch(exep) {
    createSpreadsheet = true;
  }   
  
  var inputTab;
  if(!ss) {
    var templateSpreadsheet = SpreadsheetApp.openByUrl(TEMPLATE_URL);
    var ss = templateSpreadsheet.copy(AdWordsApp.currentAccount().getName() + ' - E-Commerce Performance Report');
    SETTINGS.URL = ss.getUrl();
    log('Report Url: '+SETTINGS.URL);   
    inputTab = SpreadsheetApp.openByUrl(SETTINGS.DASHBOARD_URL).getSheetByName(INPUT_TAB_NAME);
    inputTab.getRange(SETTINGS.ROW_NUM,SETTINGS.COL_NUM-1,1,1).setValue(SETTINGS.URL);
  }
  
  var fileName = ss.getName();
  addToFolder(fileName);
  
  SETTINGS.IDS = SETTINGS.PROFILE_IDS ? SETTINGS.PROFILE_IDS.toString().split(',') : [];
  if(SETTINGS.IDS.length == 0) { return; }
  
  compileOverallSummaryReport(SETTINGS);
  compileShoppingReport(SETTINGS);
  
  if(!inputTab) {
    inputTab = SpreadsheetApp.openByUrl(SETTINGS.DASHBOARD_URL).getSheetByName(INPUT_TAB_NAME);
  }
  
  var NOW = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss');
  inputTab.getRange(SETTINGS.ROW_NUM,SETTINGS.COL_NUM,1,1).setValue(NOW);
}

function compileShoppingReport(SETTINGS) {
  var outputTab = SpreadsheetApp.openByUrl(SETTINGS.URL).getSheetByName('Google Shopping');
  outputTab.clearContents();
  
  var output = [['Top 10 AdGroups By Spend','',''],['Campaign','AdGroup','Spends']];
  var iter = AdWordsApp.shoppingAdGroups().orderBy('Cost DESC').forDateRange('LAST_30_DAYS').withLimit(10).get();
  while(iter.hasNext()) {
    var entity = iter.next();
    output.push([entity.getCampaign().getName(), entity.getName(), entity.getStatsFor('LAST_30_DAYS').getCost()]);
  }
  
  outputTab.getRange(1,1,output.length,output[0].length).setValues(output);
  
  var output = [['Top 10 AdGroups By Conversions','',''],['Campaign','AdGroup','Conversions']];
  var iter = AdWordsApp.shoppingAdGroups().orderBy('Conversions DESC').withLimit(10).forDateRange('LAST_30_DAYS').get();
  while(iter.hasNext()) {
    var entity = iter.next();
    output.push([entity.getCampaign().getName(), entity.getName(), entity.getStatsFor('LAST_30_DAYS').getConversions()]);
  }
  
  outputTab.getRange(1,6,output.length,output[0].length).setValues(output);
  
  var output = [['AdGroups without spend',''],['Campaign','AdGroup']];
  var iter = AdWordsApp.shoppingAdGroups()
  .withCondition('Status = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('Cost = 0')
  .withLimit(10)
  .forDateRange('LAST_30_DAYS').get();
  
  while(iter.hasNext()) {
    var entity = iter.next();
    output.push([entity.getCampaign().getName(), entity.getName()]);
  }
  
  outputTab.getRange(1,11,output.length,output[0].length).setValues(output);
  
  
  /*********** Shopping Search Queries ************************/
  var shoppingCampaignIds = fetchShoppingCampaignIds();
  
  if(shoppingCampaignIds.length) {
    fetchShoppingSearchQueries(shoppingCampaignIds, outputTab);
  }
  
  fetchProductStats(outputTab);
}

function fetchShoppingSearchQueries(ids, outputTab) {
  var DATE_RANGE = getAdWordsFormattedDate(30, 'yyyyMMdd') + ',' + getAdWordsFormattedDate(0,'yyyyMMdd');
  
  var zeroConversionQueries = {};
  var map = {};
  var OPTIONS = { includeZeroImpressions : false };
  
  var cols = ['CampaignName','AdGroupName','Query','Impressions','Clicks','Cost','Conversions','Ctr'];
  var reportName = 'SEARCH_QUERY_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where CampaignId IN [' + ids.join(',') + ']',
               'and AdGroupStatus = ENABLED',
               'and Clicks >= 10',
               'during',DATE_RANGE].join(' ');
  
  //Logger.log(query);
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    
    row.Ctr = parseFloat(row.Ctr.toString().replace(/%/g,''));
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    
    var key = [row.CampaignName,row.AdGroupName,row.Query].join('-');
    map[key] = row;
    if(row.Conversions == 0) {
      zeroConversionQueries[key] = row;
    }
  }
  
  var output = [['Top 10 Keywords By Spend with 0 conversions','','',''],['Campaign','AdGroup','Keyword','Spend']];
  var keysSorted = Object.keys(zeroConversionQueries).sort(function(a,b){return zeroConversionQueries[b].Cost-zeroConversionQueries[a].Cost})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = zeroConversionQueries[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.Query, row.Cost]);
  }
  
  outputTab.getRange(14,1,output.length,output[0].length).setValues(output);
  
  
  var output = [['Top 10 Keywords with Lowest CTR','','',''],['Campaign','AdGroup','Keyword','Ctr']];
  var keysSorted = Object.keys(map).sort(function(a,b){return map[a].Ctr-map[b].Ctr})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = map[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.Query, row.Ctr+'%']);
  }
  outputTab.getRange(14,6,output.length,output[0].length).setValues(output); 
}

function fetchProductStats(outputTab) {
  var DATE_RANGE = getAdWordsFormattedDate(30, 'yyyyMMdd') + ',' + getAdWordsFormattedDate(0,'yyyyMMdd');
  
  var zeroClickProducts = {};
  var map = {}, cpaMap = {}, roasMap = {}, spendsMap = {};
  var OPTIONS = { 'includeZeroImpressions' : false };
  
  var cols = ['CampaignName','AdGroupName','ProductGroup','Impressions','Clicks','Cost',
              'Conversions','Ctr','CostPerConversion','ConversionValue'];
  var reportName = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where AdGroupStatus = ENABLED',
               'during',DATE_RANGE].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.ProductGroup == '* /') { continue; }
    
    row.Clicks = parseInt(row.Clicks.toString().replace(/,/g,''),10);
    row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.CostPerConversion = parseFloat(row.CostPerConversion.toString().replace(/,/g,''));
    row.ROAS = row.Cost == 0 ? 0 : round((row.ConversionValue / row.Cost),2);
    
    var key = [row.CampaignName,row.AdGroupName,row.ProductGroup].join('-');
    map[key] = row;
    
    if(row.Cost > 0) {
      spendsMap[key] = row;
      roasMap[key] = row;
    }
    
    if(row.CostPerConversion > 0) {
      cpaMap[key] = row;
    }
    
    if(row.Clicks == 0) {
      zeroClickProducts[key] = row;
    }
  }
  
  var output = [['Products with Lowest CPA','','',''],['Campaign','AdGroup','Product','CPA']];
  var keysSorted = Object.keys(cpaMap).sort(function(a,b){return cpaMap[a].CostPerConversion-cpaMap[b].CostPerConversion})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = cpaMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.CostPerConversion]);
  }
  
  outputTab.getRange(27,1,output.length,output[0].length).setValues(output);
  
  
  var output = [['Products with Highest CPA','','',''],['Campaign','AdGroup','Product','CPA']];
  var keysSorted = Object.keys(cpaMap).sort(function(a,b){return cpaMap[b].CostPerConversion-cpaMap[a].CostPerConversion})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = cpaMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.CostPerConversion]);
  }
  outputTab.getRange(27,6,output.length,output[0].length).setValues(output); 
  
  
  var output = [['Products with Lowest ROAS','','',''],['Campaign','AdGroup','Product','ROAS']];
  var keysSorted = Object.keys(roasMap).sort(function(a,b){return roasMap[a].ROAS-roasMap[b].ROAS})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = roasMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.ROAS]);
  }
  
  outputTab.getRange(40,1,output.length,output[0].length).setValues(output);
  
  var output = [['Products with Highest ROAS','','',''],['Campaign','AdGroup','Product','ROAS']];
  var keysSorted = Object.keys(roasMap).sort(function(a,b){return roasMap[b].ROAS-roasMap[a].ROAS})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = roasMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.ROAS]);
  }
  
  outputTab.getRange(40,6,output.length,output[0].length).setValues(output);
  
  
  var output = [['Products with Lowest Spend','','',''],['Campaign','AdGroup','Product','Spend']];
  var keysSorted = Object.keys(spendsMap).sort(function(a,b){return spendsMap[a].Cost-spendsMap[b].Cost})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = spendsMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.Cost]);
  }
  
  outputTab.getRange(53,1,output.length,output[0].length).setValues(output);
  
  var output = [['Products with Highest Spend','','',''],['Campaign','AdGroup','Product','Spend']];
  var keysSorted = Object.keys(spendsMap).sort(function(a,b){return spendsMap[b].Cost-spendsMap[a].Cost})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = spendsMap[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.Cost]);
  }
  
  outputTab.getRange(53,6,output.length,output[0].length).setValues(output);
  
  var output = [['Products with Highest Clicks','','',''],['Campaign','AdGroup','Product','Clicks']];
  var keysSorted = Object.keys(map).sort(function(a,b){return map[b].Clicks-map[a].Clicks})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = map[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup, row.Clicks]);
  }
  
  outputTab.getRange(66,1,output.length,output[0].length).setValues(output);
  
  var output = [['Products with No Clicks','',''],['Campaign','AdGroup','Product']];
  var keysSorted = Object.keys(zeroClickProducts).sort(function(a,b){return zeroClickProducts[b].Cost-zeroClickProducts[a].Cost})
  for(var k = 0; k < 10; k++) {
    if(!keysSorted[k]) { break;}
    var row = zeroClickProducts[keysSorted[k]];
    output.push([row.CampaignName, row.AdGroupName, row.ProductGroup]);
  }
  
  outputTab.getRange(66,6,output.length,output[0].length).setValues(output);
  
}

function fetchShoppingCampaignIds() {
  var ids = [];
  var iter = AdWordsApp.shoppingCampaigns()
  .withCondition('Impressions > 0')
  .forDateRange(getAdWordsFormattedDate(30, 'yyyyMMdd'), getAdWordsFormattedDate(0,'yyyyMMdd'))
  .get();
  
  while(iter.hasNext()) {
    ids.push(iter.next().getId());
  }
  
  return ids;
}

function compileOverallSummaryReport(SETTINGS) {
  
  var currentMonth = getAdWordsFormattedDate(0, 'MMM');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var LAST_YEAR = parseInt(TO.substring(0,4),10)-1; 
  var FROM = LAST_YEAR + '-01-01';
  
  var years = [LAST_YEAR, TO.substring(0,4)];
  
  var stats = initStatsMap(years);
  for(var k in SETTINGS.IDS) {
    getAnalyticsStatsByMonth(SETTINGS.IDS[k],FROM,TO,stats);
  }
  
  var output = [];
  
  var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  
  for(var k in monthNames) {
    var rowNum = parseInt(k,10)+3;
    var key = monthNames[k] + ' ' + years[0];
    /*var roasOld = stats[key].cost == 0 : 0 : round((stats[key].revenue / stats[key].cost),2);
    var cpaOld = stats[key].conversions == 0 : 0 : round((stats[key].cost / stats[key].conversions),2);
    var aovOld = stats[key].conversions == 0 : 0 : round((stats[key].revenue / stats[key].conversions),2);*/
    
    var nowKey = monthNames[k] + ' ' + years[1];
    /*var roasOld = stats[nowKey].cost == 0 : 0 : round((stats[nowKey].revenue / stats[nowKey].cost),2);
    var cpaOld = stats[nowKey].conversions == 0 : 0 : round((stats[nowKey].cost / stats[nowKey].conversions),2);
    var aovOld = stats[nowKey].conversions == 0 : 0 : round((stats[nowKey].revenue / stats[nowKey].conversions),2);*/
    
    output.push([stats[key].cost, stats[nowKey].cost, '=IF(B'+rowNum+'="","",IF(B'+rowNum+'=0,1,(C'+rowNum+'-B'+rowNum+')/B'+rowNum+'))',
                 stats[key].conversions, stats[nowKey].conversions, '=IF(E'+rowNum+'="","",IF(E'+rowNum+'=0,1,(F'+rowNum+'-E'+rowNum+')/E'+rowNum+'))',
                stats[key].revenue, stats[nowKey].revenue, '=IF(H'+rowNum+'="","",IF(H'+rowNum+'=0,1,(I'+rowNum+'-H'+rowNum+')/H'+rowNum+'))',
                ]);
    if(monthNames[k] == currentMonth) { break;}    
  }
  
  var yearHeader = [LAST_YEAR, LAST_YEAR+1];
  var outputTab = SpreadsheetApp.openByUrl(SETTINGS.URL).getSheetByName('Overall');
  outputTab.getRange(2,2,1,2).setValues([yearHeader]);
  outputTab.getRange(2,5,1,2).setValues([yearHeader]);
  outputTab.getRange(2,8,1,2).setValues([yearHeader]);
  outputTab.getRange(2,11,1,2).setValues([yearHeader]);
  outputTab.getRange(2,14,1,2).setValues([yearHeader]);
  outputTab.getRange(2,17,1,2).setValues([yearHeader]);
  outputTab.getRange(2,20,1,2).setValues([yearHeader]);
  
  outputTab.getRange(19,14,1,2).setValues([yearHeader]);
  outputTab.getRange(19,16,1,2).setValues([yearHeader]);
  outputTab.getRange(19,18,1,2).setValues([yearHeader]);
  
  var lr = outputTab.getRange(3,2,output.length,output[0].length).setValues(output).getLastRow();
  if(lr > 3) {
    outputTab.getRange(lr-1, 1, 1, outputTab.getLastColumn()).setBackground('#fff');
  }
  
  outputTab.getRange(lr, 1, 1, outputTab.getLastColumn()).setBackground('#ffe599');
  
  var sourceMediumStats = {};
  var FROM = TO.substring(0,8) + '01';
  
  
  for(var k in SETTINGS.IDS) {
    getStatsBySourceMedium(SETTINGS.IDS[k],FROM,TO,sourceMediumStats);
  }
  
  var sortedKeys = Object.keys(sourceMediumStats).sort(function(a,b){return sourceMediumStats[b].revenue-sourceMediumStats[a].revenue})
  
  var output = [];
  var otherRevenue = 0;
  for(var i in sortedKeys) {
    var key = sortedKeys[i];
    if(output.length < 10) {
      output.push([key, sourceMediumStats[key].revenue]); 
      continue;
    }
    
    otherRevenue += sourceMediumStats[key].revenue;
  }
  
  if(otherRevenue > 0) {
    output.push(['Other', otherRevenue]); 
  }
  
  if(output.length > 0) {
    outputTab.getRange(19,2,output.length,output[0].length).setValues(output).sort({'column': 3, 'ascending': false});
  }
  
  
  var stats = { 'TM': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }, 
               '3M': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }, 
               '12M': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }
              }
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  date.setHours(12);
  date.setDate(0);
  
  var LM_END = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  date.setMonth(date.getMonth()-3);
  date.setDate(0);
  var FROM_3M = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  date.setMonth(date.getMonth()-9);
  date.setDate(0);
  var FROM_12M = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  var DATE_RANGE = { 'TM': { 'FROM': FROM, 'TO': TO },
                    '3M': { 'FROM': FROM_3M, 'TO': LM_END },
                    '12M': { 'FROM': FROM_12M, 'TO': LM_END }
                   }
  
  for(var k in SETTINGS.IDS) {
    getSpecificStats(SETTINGS.IDS[k],DATE_RANGE,stats);
  }
  
  
  var oldStats = { 'TM': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }, 
                  '3M': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }, 
                  '12M': { 'spend': 0, 'revenue': 0, 'adRevenue': 0 }
                 }
  
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  date.setDate(date.getDate()-1);
  date.setYear(date.getYear()-1);
  
  var TO = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  
  date.setDate(0);
  var LM_END = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  date.setMonth(date.getMonth()-3);
  date.setDate(0);
  var FROM_3M = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  date.setMonth(date.getMonth()-9);
  date.setDate(0);
  var FROM_12M = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  
  var DATE_RANGE_OLD = { 'TM': { 'FROM': FROM, 'TO': TO },
                        '3M': { 'FROM': FROM_3M, 'TO': LM_END },
                        '12M': { 'FROM': FROM_12M, 'TO': LM_END }
                       }
  
  for(var k in SETTINGS.IDS) {
    getSpecificStats(SETTINGS.IDS[k],DATE_RANGE_OLD,oldStats);
  }
  
  var output = [];
  for(var key in stats) {
    output.push([oldStats[key].spend, stats[key].spend, oldStats[key].revenue, stats[key].revenue]);
  }
  
  if(output.length > 0) {
    outputTab.getRange(20,14,output.length,output[0].length).setValues(output);
  }
}


function getLinkedProfileIds(UA_NUMBER) {  
  var data = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL).getSheetByName(ANALYTICS_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  UA_NUMBER = UA_NUMBER.toString();
  var profiles = [];
  for(var k in data) {
    if(UA_NUMBER.indexOf(data[k][2]) > -1) {
      profiles.push(data[k][3]);
    }
  }
  
  return profiles;
}

function getSpecificStats(id, DATE_RANGE, stats) {
  var optArgs = { 'filters': 'ga:source==google;ga:medium==cpc' }
  for(var key in DATE_RANGE) {
    var attempts = 3;
    var results;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        results = Analytics.Data.Ga.get(
          'ga:'+id,              // Table id (format ga:xxxxxx).
          DATE_RANGE[key]['FROM'],                 // Start-date (format yyyy-MM-dd).
          DATE_RANGE[key]['TO'],                   // End-date (format yyyy-MM-dd).
          "ga:adCost,ga:transactionRevenue",
          optArgs);
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + id);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var rows = results.getRows();
    
    for(var k in rows) {
      stats[key].spend += parseFloat(rows[k][0]);
      stats[key].adRevenue += parseFloat(rows[k][1]);
    }
  }
  
  
  for(var key in DATE_RANGE) {
    var attempts = 3;
    var results;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        results = Analytics.Data.Ga.get(
          'ga:'+id,              // Table id (format ga:xxxxxx).
          DATE_RANGE[key]['FROM'],                 // Start-date (format yyyy-MM-dd).
          DATE_RANGE[key]['TO'],                   // End-date (format yyyy-MM-dd).
          "ga:transactionRevenue");
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + id);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var rows = results.getRows();
    
    for(var k in rows) {
      stats[key].revenue += parseFloat(rows[k][0]);
    }
  }
}


function getAnalyticsStatsByMonth(id,FROM,TO,stats) {
  var attempts = 3;
  var results;
  
  var optArgs = { 'dimensions': 'ga:yearMonth', 'filters': 'ga:source==google;ga:medium==cpc' }
  // Make a request to the API.
  while(attempts > 0) {
    try {
      results = Analytics.Data.Ga.get(
        'ga:'+id,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = results.getRows();
  
  var monthNames = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  for(var k in rows) {
    var key = monthNames[parseInt(rows[k][0].toString().substring(4,6),10)] + ' ' + rows[k][0].toString().substring(0,4);
    if(!stats[key]) { continue; }
    stats[key].cost += parseFloat(rows[k][1]);
    stats[key].conversions += parseInt(rows[k][2],10);    
    stats[key].revenue += parseFloat(rows[k][3]);
  }
}

function getStatsBySourceMedium(id,FROM,TO,stats) {
  var attempts = 3;
  var results;
  
  var optArgs = { 'dimensions': 'ga:sourceMedium', 'filters': 'ga:transactionRevenue>0', 'sort': '-ga:transactionRevenue' }
  // Make a request to the API.
  while(attempts > 0) {
    try {
      results = Analytics.Data.Ga.get(
        'ga:'+id,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = results.getRows();
  for(var k in rows) {
    if(!stats[rows[k][0]]) { stats[rows[k][0]] = { 'revenue': 0 } }
    stats[rows[k][0]].revenue += parseFloat(rows[k][1]);
  }
}

function initStatsMap(years) {
  var map = {};
  
  var monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  
  for(var k in years) {
    for(var j in monthNames) {
      var key = monthNames[j] + ' ' + years[k];
      map[key] = { 'cost':0, 'conversions': 0, 'revenue':0, 'roas': 0, 'cpa': 0, 'aov': 0 }
    }
  }
  return map;
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
} 


function addToFolder(fileName) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  
  var fileIter = DriveApp.getRootFolder().searchFiles("title contains '" + fileName + "'");
  while(fileIter.hasNext()){
    var file = fileIter.next();
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }     
}

function log(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg); 
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}