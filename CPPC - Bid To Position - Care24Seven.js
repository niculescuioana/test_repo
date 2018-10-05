/***********************************************
* Bid To Position - Care24Seven
* @author: Naman Jindal (nj.itprof@gmail.com)
* @version: 1.0
************************************************/

var SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/1VakpcQc5vFTagvJAGY7eKwdXfQHagDG8tNphgce2Nb0/edit#gid=0';

function main() {
  MccApp.accounts().withCondition('Name IN ["C247 - Chiswick", "C247 - Guilford", "C247 - Windsor"]').executeInParallel('run')
}

function run() {
  var SETTINGS = parseInputs();
  if(SETTINGS.FLAG.toLowerCase() != 'yes') { return; }
  
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'),10);
  //Logger.log(hour);
  //Logger.log(SETTINGS.HOURS_BETWEEN);
  
  SETTINGS.HOURS_BETWEEN = SETTINGS.HOURS_BETWEEN.split('-');
  SETTINGS.HOURS_BETWEEN[0] = parseInt(SETTINGS.HOURS_BETWEEN[0], 10);
  SETTINGS.HOURS_BETWEEN[1] = parseInt(SETTINGS.HOURS_BETWEEN[1], 10);

  if(hour < SETTINGS.HOURS_BETWEEN[0] || hour > SETTINGS.HOURS_BETWEEN[1]) { 
    Logger.log('Not Scheduled to run now');  
    return; 
  }
  
  var ids = [];
  if(SETTINGS.CAMPAIGN_LABEL_IS_NOT) {
   var camps = AdWordsApp.campaigns().withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.CAMPAIGN_LABEL_IS_NOT + '"]').get();
    while(camps.hasNext()) {
     ids.push(camps.next().getId());
    }
  }
  
  
  var iter = AdWordsApp.keywords()
  .withCondition('Status = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .withCondition('Impressions > 0')
  .withCondition('AveragePosition > 0')
  .withCondition('AveragePosition < ' + SETTINGS.TARGET_POS_MAX)
  .withCondition('MaxCpc > ' + SETTINGS.MIN_CPC)
  .forDateRange('TODAY');
  
  if(SETTINGS.KEYWORD_LABEL_IS) {
   iter.withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.KEYWORD_LABEL_IS + '"]') 
  }
  
  if(ids.length) {
    iter.withCondition('CampaignId NOT_IN [' + ids.join(',') + ']');
  }
  
  iter = iter.get()
  
  while(iter.hasNext()) {
    var kw = iter.next();
    var newCpc = kw.bidding().getCpc()*(1-SETTINGS.DECREASE_FACTOR);
    if(newCpc < SETTINGS.MIN_CPC) { newCpc = SETTINGS.MIN_CPC; }
    kw.bidding().setCpc(newCpc);
  }
  
  
  var iter = AdWordsApp.keywords()
  .withCondition('Status = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .withCondition('Impressions > 0')
  .withCondition('AveragePosition > ' + SETTINGS.TARGET_POS_MIN)
  .withCondition('MaxCpc < ' + SETTINGS.MAX_CPC)
  .forDateRange('TODAY');
  
  if(SETTINGS.KEYWORD_LABEL_IS) {
   iter.withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.KEYWORD_LABEL_IS + '"]') 
  }
  
  if(ids.length) {
    iter.withCondition('CampaignId NOT_IN [' + ids.join(',') + ']');
  }
  
  iter = iter.get();
  while(iter.hasNext()) {
    var kw = iter.next();
    var newCpc = kw.bidding().getCpc()*(1+SETTINGS.INCREASE_FACTOR);
    if(newCpc > SETTINGS.MAX_CPC) { newCpc = SETTINGS.MAX_CPC; }
    kw.bidding().setCpc(newCpc);
  }
 
  //return;
  var date = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM, dd yyyy HH:mm');
  SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName('Settings').getRange(12, SETTINGS.COL).setValue(date)
}

function parseInputs() {
  var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName('Settings').getDataRange().getValues();
  var accounts = data.shift();
  var index = accounts.indexOf(AdWordsApp.currentAccount().getName());
  
  var HEADER = [
    'FLAG', 'CAMPAIGN_LABEL_IS_NOT', 'KEYWORD_LABEL_IS',
    'HOURS_BETWEEN', 'MAX_CPC', 'MIN_CPC', 
    'TARGET_POS_MAX', 'TARGET_POS_MIN', 'INCREASE_FACTOR', 'DECREASE_FACTOR'
  ];
  
  var SETTINGS = {'COL': index+1};
  for(var k in HEADER) {
    SETTINGS[HEADER[k]] = data[k][index]; 
  }
  
  return SETTINGS;
}