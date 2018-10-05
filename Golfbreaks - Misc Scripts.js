/******************************************
* Combination of multiple scripts:
*
* AdGroup Export - Daily 8 AM
* Extensions Overview - Monday 4 AM
* Copy Keyword Urls - Hourly
*
*
* @author: Naman Jindal (naman@pushgroup.co.uk)
* @version: 2.0
******************************************/

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1frtLJbPbtV4_9v0kfOtd_P65cwYpk8lWkz7EdznyD1M/edit#gid=0';

function main() { 
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  if(now.getHours() == 8) {
     runAdGroupExportScript();
  } else if(now.getDay() == 1 && now.getHours() == 4) {
    runExtensionsOverviewScript();
  } else {
    //runCopyUrlScript();
  }
}

function runCopyUrlScript() {
  var HEADER = ['ID', 'NAME', 'FLAG'];
  var accountMap = scanAccounts('Copy Keyword Urls', HEADER);
  var ids = Object.keys(accountMap);
  
  MccApp.accounts()
  .withIds(ids)
  .executeInParallel('copyKeywordUrls', 'callBack', JSON.stringify(accountMap));
}

function runAdGroupExportScript() {
  var HEADER = ['ID', 'NAME', 'FLAG', 'URL', 'SHEET_NAME'];
  var accountMap = scanAccounts('AdGroup Export', HEADER);
  var ids = Object.keys(accountMap);
  
  MccApp.accounts()
  .withIds(ids)
  .executeInParallel('exportAdGroups', 'callBack', JSON.stringify(accountMap));
}

function runExtensionsOverviewScript() {
  var HEADER = ['ID', 'NAME', 'FLAG', 'URL', 'SHEET_NAME'];
  var accountMap = scanAccounts('Extensions Overview', HEADER);
  var ids = Object.keys(accountMap);
  
  MccApp.accounts()
  .withIds(ids)
  .executeInParallel('compileExtensionsOverview', 'callBack', JSON.stringify(accountMap));
}

function callBack() {
  
}

function scanAccounts(TAB_NAME, HEADER) {
  var map = {};
  var data = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  data.shift();
  
  
  for(var k in data) {
    if(!data[k][0] || data[k][2] != 'Yes') { continue; }
    var SETTINGS = {};
    for(var j in HEADER) {
      SETTINGS[HEADER[j]] = data[k][j]; 
    }
    map[data[k][0]] = SETTINGS;
  }
  
  return map;
}

function copyKeywordUrls(input) {
  
  var iter = AdWordsApp.keywords()
  .withCondition('Status = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('FinalUrls DOES_NOT_CONTAIN_IGNORE_CASE "h"')
  .get();
  
  if(iter.totalNumEntities() == 0) { return; }
  
  var map = {};
  while(iter.hasNext()) {
   var kw = iter.next(); 
    var agId = kw.getAdGroup().getId();
    if(!map[agId]) { map[agId] = []; }
    
    map[agId].push(kw);
  }
  
  var agIds = Object.keys(map);
  
  var OPTIONS = { 'includeZeroImpressions' : true };
  var cols = ['AdGroupId','Id','FinalUrls','Impressions','Clicks'];
  var reportName = 'KEYWORDS_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where Status = ENABLED',
               'and AdGroupStatus = ENABLED',
               'and CampaignStatus = ENABLED',
               'and AdGroupId IN [' + agIds.join(',') + ']',
               'and FinalUrls CONTAINS_IGNORE_CASE "h"',
               'during','TODAY'].join(' ');
  
  var urlMap = {};
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(!row.FinalUrls) { continue; }
    
    urlMap[row.AdGroupId] = row.FinalUrls;
  }
  
  for(var agId in map) {
    if(!urlMap[agId]) { continue; }
    for(var i in map[agId]) {
      map[agId][i].urls().setFinalUrl(urlMap[agId]);
    }
  }
}

function exportAdGroups(input) {
  
  var SETTINGS = JSON.parse(input)[AdWordsApp.currentAccount().getCustomerId()];
  
  var output = [['Campaign', 'AdGroup']];
  var adGroups = AdWordsApp.adGroups()
  .withCondition('Status = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .orderBy('CampaignName ASC')  
  .get();
  
  while(adGroups.hasNext()) {
    var ag = adGroups.next();
    var agName = ag.getName();
    var campName = ag.getCampaign().getName();
    
    output.push([campName, agName]);
  }
  
  
  var ss = SpreadsheetApp.openByUrl(SETTINGS.URL);
  var tab = ss.getSheetByName(SETTINGS.SHEET_NAME);
  if(!tab) {
    tab = ss.insertSheet(SETTINGS.SHEET_NAME);
  }
  
  tab.setFrozenRows(1);
  tab.clearContents();
  tab.getRange(1, 1, output.length, output[0].length).setValues(output);
  tab.getRange(1, 1, 1, output[0].length).setBackground('#efefef').setFontWeight('bold');
  tab.getRange(2, 1, output.length-1, output[0].length).sort([{'column': 1}, {'column': 2}])
  tab.getDataRange().setFontFamily('Calibri');
}

function compileExtensionsOverview(input) {
  
  var SETTINGS = JSON.parse(input)[AdWordsApp.currentAccount().getCustomerId()];
  
  var output = [['Campaign', 'AdGroup', 'Num Sitelink Extensions', 'Num Callout Extensions', 'Call Extensions']];
  var campaigns = AdWordsApp.campaigns()
  .withCondition('Impressions > 0')
  //.withCondition('AdNetworkType1 = SEARCH')
  .forDateRange('LAST_14_DAYS')
  .get();
  
  while(campaigns.hasNext()) {
    var camp = campaigns.next();
    var campName = camp.getName();
    
    output.push([campName, '', camp.extensions().sitelinks().get().totalNumEntities(),
                 camp.extensions().callouts().get().totalNumEntities(), 
                 camp.extensions().phoneNumbers().get().totalNumEntities()]);
  }
  
  var adGroups = AdWordsApp.adGroups()
  .withCondition('Impressions > 0')
  //.withCondition('AdNetworkType1 = SEARCH')
  .forDateRange('LAST_14_DAYS')
  .get();
  
  while(adGroups.hasNext()) {
    var ag = adGroups.next();
    var agName = ag.getName();
    var campName = ag.getCampaign().getName();
    
    output.push([campName, agName, ag.extensions().sitelinks().get().totalNumEntities(),
                 ag.extensions().callouts().get().totalNumEntities(), 
                 ag.extensions().phoneNumbers().get().totalNumEntities()]);
  }
  
  
  var tab = SpreadsheetApp.openByUrl(SETTINGS.URL).getSheetByName(SETTINGS.SHEET_NAME);
  tab.clearContents();
  tab.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function info(msg) {
 Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg); 
}