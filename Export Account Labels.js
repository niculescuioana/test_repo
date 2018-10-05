
var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=153054321';
var TAB_NAME = 'Account Label Manager';


var MCC_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1F3bjn411jR3aYEJpLNAeCdqobJYlMQRUlK-6KVcdjCE/edit';
var RAW_REPORT_TAB_NAME = 'Daily Report';
var CURRENCY_EXCHANGE_TAB_NAME = 'Currency Exchange';



function main() {
  var hour = parseInt(Utilities.formatDate(new Date(), 'GMT', 'HH'), 10);
  if(hour == 7) {
    recordDailySpends(1); 
  }
  
  //return;
  var accountMap = readAccounts();
  var out = [];
  for(var id in accountMap) {
    var account = accountMap[id];
    if(!account) { continue; }
    
    var map = {};
    var labelIter = account.labels().get();
    while(labelIter.hasNext()) {
      map[labelIter.next().getName()] = 1 ;
    }
    
    var existing = Object.keys(map);
    out.push([id, existing.join('; '), account.getStatsFor('LAST_WEEK').getClicks()]);
  }
  
  
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME) ;
  tab.getRange(2, 1, out.length, out[0].length).setValues(out);
  
  //exportLabels();
}

function readAccounts() {
  var map = {}
  
  var accounts = MccApp.accounts().get();
  while(accounts.hasNext()) {
    var account = accounts.next();
    MccApp.select(account);
    map[account.getCustomerId()] = account;
  }
  
  return map;
}

function createLabelIfNeeded(label) {
  try {
    var labelIter = MccApp.accountLabels()
    .withCondition("LabelNames CONTAINS '"+label+"'")
    .get();
  } catch(e) {
    MccApp.createAccountLabel(label);
  }
}


function recordDailySpends(N) {
  var date = getAdWordsFormattedDate(N, 'MMM dd, yyyy');
  
  var dateRange = getAdWordsFormattedDate(N, 'yyyyMMdd')
  var mccAccount = AdWordsApp.currentAccount();
  
  /*var spendMCC = mccAccount.getStatsFor(dateRange,dateRange).getCost();
  //Logger.log(spendMCC);
  //return;
  var ignoreSpend = 0;
  
  var CURRENCY_MAP = readCurrencyExchangeRates();
  var map = readAccountsToIgnore();
  var ids = Object.keys(map);
  
  var iter = MccApp.accounts()
  .withIds(ids)
  .withCondition('Cost > 0')
  .forDateRange(dateRange,dateRange)
  .get();
  
  
  while(iter.hasNext()) {
    MccApp.select(iter.next());
    if(!AdWordsApp.currentAccount().getName()) { continue; }
    
    var spends = AdWordsApp.currentAccount().getStatsFor(dateRange, dateRange).getCost();
    var conversions = AdWordsApp.currentAccount().getStatsFor(dateRange, dateRange).getConversions();    
    
    var currency = AdWordsApp.currentAccount().getCurrencyCode();
    
    var gbpSpends = spends;
    if(currency != 'GBP') {
      gbpSpends = spends*CURRENCY_MAP[currency];
    }
    
    ignoreSpend += gbpSpends;
  }*/
  
  
  var spendMCC = compileYesterdaySummary(N);
  var tab = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName('Daily Summary');
  tab.appendRow([date, '', '', spendMCC]);
  tab.sort(1, false);
  
  MccApp.select(mccAccount);
}

function compileYesterdaySummary(N) {
  var date = getAdWordsFormattedDate(N, 'MMM dd, yyyy');
  
  var dateRange = getAdWordsFormattedDate(N, 'yyyyMMdd')
  var mccAccount = AdWordsApp.currentAccount();
  
  //var spendMCC = mccAccount.getStatsFor(dateRange,dateRange).getCost();
  //Logger.log(spendMCC);
  //return;
  //var ignoreSpend = 0;
  
  var CURRENCY_MAP = readCurrencyExchangeRates();
  var map = readAccountsToIgnore();
  var ids = Object.keys(map);
  
  var out = [];
  var iter = MccApp.accounts()
  //.withIds(ids)
  .withCondition('Cost > 0')
  .forDateRange(dateRange,dateRange)
  .get();
  
  
  while(iter.hasNext()) {
    MccApp.select(iter.next());
    if(!AdWordsApp.currentAccount().getName()) { continue; }
    if(ids.indexOf(AdWordsApp.currentAccount().getCustomerId()) > -1) { continue; }
    
    var spends = AdWordsApp.currentAccount().getStatsFor(dateRange, dateRange).getCost();
    var conversions = AdWordsApp.currentAccount().getStatsFor(dateRange, dateRange).getConversions();    
    
    var currency = AdWordsApp.currentAccount().getCurrencyCode();
    
    var gbpSpends = spends;
    if(currency != 'GBP') {
      gbpSpends = spends*CURRENCY_MAP[currency];
    }
    
    out.push([AdWordsApp.currentAccount().getName(), AdWordsApp.currentAccount().getCustomerId(),
              AdWordsApp.currentAccount().getCurrencyCode(), spends, gbpSpends]);
  }
  
  
  var tab = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName('Yesterday Summary');
  tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  tab.getRange(3,1,out.length,out[0].length).setValues(out);
  tab.sort(5, false);
  
  return tab.getRange('E2').getValue();
  
  //MccApp.select(mccAccount);
}

function readAccountsToIgnore() {
  var data = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName('Ignore Accounts').getDataRange().getValues();
  data.shift();
  
  var map = {};
  for(var x in data) {
    if(data[x][0]) {
      map[data[x][0]] = 1;
    }
    
    if(data[x][1]) {
      map[data[x][1]] = 1;
    }    
  }
  
  return map;
}

function readCurrencyExchangeRates() {
  var map = {};
  var data = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName(CURRENCY_EXCHANGE_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][0]) { continue; }
    map[data[k][0]] = data[k][1]; 
  }
  
  return map;
}



function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}
