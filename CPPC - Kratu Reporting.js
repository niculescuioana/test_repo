var REPORT_URL = 'https://docs.google.com/spreadsheets/d/12wsp5DYyRgVANU_2pQ5xLd4CLnth3WtGJdp6jS-aao8/edit#gid=1203516540';
var TAB_NAME = 'Kratu Overview';


var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_TAB_NAME = 'Dashboard Urls';


function main() {
  var accountMap = {};
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME) ;
  var data = tab.getDataRange().getValues();
  data.shift();
  
  
  for(var k in data) {
    var manager = data[k][0]
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "' + manager + '"').get();
    while(accounts.hasNext()) {
      var account = accounts.next();
      if(!accountMap[account.getName()]) { 
        accountMap[account.getName()] = { 'ID': account.getCustomerId(), 'MANAGERS': [] }
        accountMap[account.getName()].MANAGERS.push(manager);
      }
    }
  }
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(TAB_NAME);
  var data = sheet.getDataRange().getValues();
  
  for(var z in data) {
    if(accountMap[data[z][1]]) {
      delete accountMap[data[z][1]];
    }
  }
  
  var out = [];
  for(var name in accountMap) {
    out.push([accountMap[name].ID, name, accountMap[name].MANAGERS.join('/')]);
  }
  
  sheet.getRange(sheet.getLastRow()+1, 1, out.length, out[0].length).setValues(out);
  
  return;
  
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(TAB_NAME);
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), "MMM dd,yyyy HH:mm:ss"));
  if(now.getHours() == 0) {
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  }
  
  if(now.getHours() < 2) {
    return;
  }
  
  var map = {};
  var data = sheet.getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][0]) { continue; }
    map[data[k][0]] = 1;  
  }
  
  var accounts = MccApp.accounts()
  //.withCondition('LabelNames CONTAINS "Yannis"')
  .withCondition('Impressions > 0')
  .forDateRange('LAST_14_DAYS')
  .get();
  
  Logger.log(accounts.totalNumEntities());
  if(accounts.totalNumEntities() == 0) { return; }
  
  while(accounts.hasNext()) {
    var account = accounts.next();
    MccApp.select(account);
    if(map[AdWordsApp.currentAccount().getCustomerId()]) { continue; }
    
    var label = '';
    var labels = accountMap[account.getName()];
    if(labels) {
      label = labels.join('/'); 
    }
    
    //Logger.log(account.getName());
    try {
      var row = runScript(label);
      sheet.appendRow(row);
    } catch(ex) {
      Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + ex); 
    }
    
    if(shouldExitNow()){
      break;
    }
  }
  
  sheet.sort(3);
}

function shouldExitNow() {
  return (AdWordsApp.getExecutionInfo().getRemainingTime() < 90) 
}

function runScript(label) {
  
  var phraseKeywords = AdWordsApp.keywords()
  .withCondition('KeywordMatchType = PHRASE')
  .withCondition('Impressions > 0')
  .forDateRange('LAST_30_DAYS')
  .withCondition('Status = ENABLED')
  .get().totalNumEntities();
  
  var exactKeywords = AdWordsApp.keywords()
  .withCondition('KeywordMatchType = EXACT')
  .withCondition('Impressions > 0')
  .forDateRange('LAST_30_DAYS')
  .withCondition('Status = ENABLED')
  .get().totalNumEntities();
  
  var broadKeywords = AdWordsApp.keywords()
  .withCondition('KeywordMatchType = BROAD')
  .withCondition('Impressions > 0')
  .forDateRange('LAST_30_DAYS')
  .withCondition('Status = ENABLED')
  .get().totalNumEntities();
  
  var totalKeywords = phraseKeywords + exactKeywords + broadKeywords;  
  
  var phrasepct = totalKeywords == 0 ? '0%' : round((100*phraseKeywords / totalKeywords),2)+'%';
  var exactpct = totalKeywords == 0 ? '0%' : round((100*exactKeywords / totalKeywords),2)+'%';
  var broadpct = totalKeywords == 0 ? '0%' : round((100*broadKeywords / totalKeywords),2)+'%';
  
  
  var ctrMap = getCtrByAdNetwork();
  var qsMap = getQSByKeywords();
  
  var camps = AdWordsApp.campaigns()
  .withCondition('Impressions > 0')
  .withCondition('AdNetworkType1 = SEARCH')
  .forDateRange('LAST_7_DAYS')
  .get();
  
  var sitelinkCamps = 0, callCamps = 0, calloutCamps = 0, reviewCamps = 0, locationCamps = 0;
  var numCampaigns = camps.totalNumEntities();
  
  
  if(numCampaigns == 0) {
    return [AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName(), 
            label, numCampaigns,0,0,0,0,0,totalKeywords,phrasepct,exactpct,broadpct, 
            ctrMap['search']['ctr'], ctrMap['display']['ctr'],
            qsMap['poorQsPct'], qsMap['avgQsPct'], qsMap['goodQsPct']];
  }
  
  var campIds = [];
  while(camps.hasNext()) {
    var camp = camps.next();
    if(camp.extensions().sitelinks().get().totalNumEntities()) {
      sitelinkCamps++; 
    }
    if(camp.extensions().phoneNumbers().get().totalNumEntities()) {
      callCamps++; 
    }
    if(camp.extensions().callouts().get().totalNumEntities()) {
      calloutCamps++; 
    }
    if(camp.extensions().reviews().get().totalNumEntities()) {
      reviewCamps++; 
    }
    
    if(camp.targeting().targetedLocations().get().totalNumEntities() || camp.targeting().targetedProximities().get().totalNumEntities()) {
      locationCamps++; 
    }
  }
  
  var sitelinkpct = round((100*sitelinkCamps / numCampaigns),2)+'%';
  var callpct = round((100*callCamps / numCampaigns),2)+'%';
  var calloutpct = round((100*calloutCamps / numCampaigns),2)+'%';
  var reviewpct = round((100*reviewCamps / numCampaigns),2)+'%';
  var locationpct = round((100*locationCamps / numCampaigns),2)+'%';
  
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status != ENABLED').get().hasNext();
  var remarketingFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  if(remarketingFlag != 'Yes') {
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'rmkt'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'rmkt'").withCondition('Status != ENABLED').get().hasNext();
    remarketingFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  }
  
  if(remarketingFlag != 'Yes') {
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 're-marketing'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 're-marketing'").withCondition('Status != ENABLED').get().hasNext();
    remarketingFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  }
  
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status != ENABLED').get().hasNext();
  var remarketingFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gsp'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gsp'").withCondition('Status != ENABLED').get().hasNext();
  var gspFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gmail'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gmail'").withCondition('Status != ENABLED').get().hasNext();
  var gmailFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'display'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'display'").withCondition('Status != ENABLED').get().hasNext();
  var displayFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'video'").withCondition('Status = ENABLED').get().hasNext();
  var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'video'").withCondition('Status != ENABLED').get().hasNext();
  var videoFlag = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  
  return [AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName(), label, numCampaigns,
          remarketingFlag, gspFlag, gmailFlag, displayFlag, videoFlag, sitelinkpct, locationpct,
          callpct, calloutpct, reviewpct, totalKeywords, phrasepct, exactpct, broadpct,
          ctrMap['search']['ctr'], ctrMap['display']['ctr'], qsMap['poorQsPct'], 
          qsMap['avgQsPct'], qsMap['goodQsPct']];
  
}

function getCtrByAdNetwork() {
  var map = { 'display': { 'clicks': 0, 'impressions': 0 }, 'search': { 'clicks': 0, 'impressions': 0 } }; 
  
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['AdNetworkType1','Impressions','Clicks'];
  var report = 'ACCOUNT_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where Clicks > 0',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.AdNetworkType1.indexOf('Search') > -1) {
      map['search']['clicks'] += parseInt(row.Clicks,10);
      map['search']['impressions'] += parseInt(row.Impressions,10);      
    } else if(row.AdNetworkType1.indexOf('Display') > -1) {
      map['display']['clicks'] += parseInt(row.Clicks,10);
      map['display']['impressions'] += parseInt(row.Impressions,10);      
    }
  }
  
  map['display']['ctr'] = map['display']['impressions'] == 0 ? '0%' : map['display']['clicks'] / map['display']['impressions'];
  map['search']['ctr'] = map['search']['impressions'] == 0 ? '0%' : map['search']['clicks'] / map['search']['impressions'];
  
  return map;
}

function getQSByKeywords() {
  var map = {};
  var qsMap = { 'total': 0, 'goodQsKeywords': 0, 'avgQsKeywords': 0, 'poorQsKeywords': 0 }
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['AdGroupId','Id','Impressions','Clicks','QualityScore'];
  var report = 'KEYWORDS_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where Status = ENABLED',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    row.QualityScore = parseInt(row.QualityScore,10);
    
    if(isNaN(row.QualityScore)) { continue; }
    
    qsMap['total']++;
    if(row.QualityScore >= 5 && row.QualityScore <= 7) {
      qsMap['avgQsKeywords']++;
    } else if(row.QualityScore >= 8 && row.QualityScore <= 10) {
      qsMap['goodQsKeywords']++;
    } else {
      qsMap['poorQsKeywords']++;
    }
  }
  
  
  qsMap['goodQsPct'] = qsMap['total'] == 0 ? '0%' : round((100*qsMap['goodQsKeywords'] / qsMap['total']),2)+'%';
  qsMap['poorQsPct'] = qsMap['total'] == 0 ? '0%' : round((100*qsMap['poorQsKeywords'] / qsMap['total']),2)+'%';
  qsMap['avgQsPct'] = qsMap['total'] == 0 ? '0%' : round((100*qsMap['avgQsKeywords'] / qsMap['total']),2)+'%';
  
  /*qsMap['avgQs'] = qsMap['Impressions'] == 0 ? 0 : qsMap['QualityScore'] / qsMap['Impressions'];
  for(var key in map) {
  if(map[key] >  qsMap['avgQs']) {
  qsMap['aboveAvgQsKeywords']++;
  }
  }*/
  
  return qsMap;
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}


function backFill() {
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(TAB_NAME);
  
  
  var map = {};
  var data = sheet.getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][0]) { continue; }
    map[data[k][0]] = parseInt(k,10);  
  }
  
  var accounts = MccApp.accounts()
  .withCondition('Impressions > 0')
  .forDateRange('LAST_14_DAYS')
  .get();
  
  if(accounts.totalNumEntities() == 0) { return; }
  
  while(accounts.hasNext()) {
    var account = accounts.next();
    MccApp.select(account);
    var rNum = map[AdWordsApp.currentAccount().getCustomerId()];
    if(!rNum && rNum !== 0) { continue; }
    
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'remarketing'").withCondition('Status != ENABLED').get().hasNext();
    data[rNum][3] = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
    
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gsp'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gsp'").withCondition('Status != ENABLED').get().hasNext();
    data[rNum][4] = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
    
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gmail'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'gmail'").withCondition('Status != ENABLED').get().hasNext();
    data[rNum][5] = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
    
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'display'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'display'").withCondition('Status != ENABLED').get().hasNext();
    data[rNum][6] = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
    
    var hasEnabled = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'video'").withCondition('Status = ENABLED').get().hasNext();
    var hasPaused = AdWordsApp.campaigns().withCondition("Name CONTAINS_IGNORE_CASE 'video'").withCondition('Status != ENABLED').get().hasNext();
    data[rNum][7] = hasEnabled ? 'Yes' : hasPaused ? 'Yes, but paused' : 'No'; 
  }
  
  sheet.getRange(2,1,data.length,data[0].length).setValues(data);
}