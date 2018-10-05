var URL = 'https://docs.google.com/spreadsheets/d/1SiIQjrC486WTB2xno_0o6RKaepYmK-g_cfpo6qL2DnI/edit?ts=599c1148';
var TAB_NAME = 'AdGroup Bid To Position';

function main() {
  var inputMap = parseSettings();
  var ids = Object.keys(inputMap);
  
  if(!ids.length) { return; }
  
  MccApp.accounts().withCondition('LabelNames DOES_NOT_CONTAIN "Not Live"').withIds(ids).executeInParallel('run', 'compile', JSON.stringify(inputMap));
}

function parseSettings() {
  var data = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME).getDataRange().getValues();
  data.shift();
  data.shift();
  
  var inputHeader = ['ACC_NAME', 'ACC_ID', 'FLAG', 'AG_LABEL', 'LAST_N_DAYS', 'POS_BETTER_THAN',
                     'DECREASE_PCT', 'MIN_BID', 'POS_WORSE_THAN', 'INCREASE_PCT', 'MAX_BID'];
  
  var map = {};
  for(var z in data) {
    if(data[z][2] != 'Y') { continue; }
    
    if(!data[z][5] || !data[z][7] || !data[z][8] || !data[z][10]) { continue; }
    
    map[data[z][1]] = { 'ROW_NUM': parseInt(z, 10)+3 }
    
    for(var k in inputHeader) {
      map[data[z][1]][inputHeader[k]] = data[z][k];
    }
  }
  
  return map;
}

function compile() {}

function run(inputMap) {
  var SETTINGS = JSON.parse(inputMap)[AdWordsApp.currentAccount().getCustomerId()];
  
  if(SETTINGS.INCREASE_PCT > 0.5 || SETTINGS.DECREASE_PCT > 0.5 ) { return; }
  
  var campIds = [];
  var camps = AdWordsApp.campaigns().withCondition('AdNetworkType1 = SEARCH').get();
  while(camps.hasNext()) {
    campIds.push(camps.next().getId()); 
  }
  
  if(campIds.length == 0) { return ''; }
  
  var END_DATE = getAdWordsFormattedDate(1, 'yyyyMMdd');
  var START_DATE = getAdWordsFormattedDate(SETTINGS.LAST_N_DAYS, 'yyyyMMdd');
  
  var ags = AdWordsApp.adGroups()
  .withCondition('Status = ENABLED')
  .withCondition('AdNetworkType1 = SEARCH')
  .withCondition('CampaignId IN [' + campIds.join(',') + ']')
  .withCondition('AveragePosition < ' + SETTINGS.POS_BETTER_THAN)
  .withCondition('KeywordMaxCpc > ' + SETTINGS.MIN_BID);
  
  if(SETTINGS.AG_LABEL) {
    ags.withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.AG_LABEL + '"]');
  }
  
  ags = ags.forDateRange(START_DATE, END_DATE).get();
  while(ags.hasNext()){
    var ag = ags.next();
    var newCpc = ag.getKeywordMaxCpc()*(1 - SETTINGS.DECREASE_PCT);
    if(ag.getKeywordMaxCpc() - newCpc < 0.01) {
      newCpc = ag.getKeywordMaxCpc() - 0.01;
    }
    
    if(newCpc < SETTINGS.MIN_BID) { newCpc = SETTINGS.MIN_BID; }
    ag.bidding().setCpc(newCpc);
  }
  
  var ags = AdWordsApp.adGroups()
  .withCondition('Status = ENABLED')
  .withCondition('AdNetworkType1 = SEARCH')
  .withCondition('CampaignId IN [' + campIds.join(',') + ']')
  .withCondition('AveragePosition > ' + SETTINGS.POS_WORSE_THAN)
  .withCondition('KeywordMaxCpc < ' + SETTINGS.MAX_BID);
  
  if(SETTINGS.AG_LABEL) {
    ags.withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.AG_LABEL + '"]');
  }
  
  ags = ags.forDateRange(START_DATE, END_DATE).get();
  while(ags.hasNext()){
    var ag = ags.next();
    var newCpc = ag.getKeywordMaxCpc()*(1 + SETTINGS.INCREASE_PCT);
    if(newCpc - ag.getKeywordMaxCpc() < 0.01) {
      newCpc = ag.getKeywordMaxCpc() + 0.01;
    }
    
    if(newCpc > SETTINGS.MAX_BID) { newCpc = SETTINGS.MAX_BID; }
    ag.bidding().setCpc(newCpc);
  }
  
  var inputSheet = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
  inputSheet.getRange(SETTINGS.ROW_NUM,12).setValue(getAdWordsFormattedDate(0, 'MMM d, yyyy'));
  
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}


function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}