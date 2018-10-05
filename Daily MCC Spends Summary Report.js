

var MCC_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1F3bjn411jR3aYEJpLNAeCdqobJYlMQRUlK-6KVcdjCE/edit#gid=112312249';
//var MCC_REPORT_URL = 'https://docs.google.com/spreadsheets/d/12rGv-WCOvtpy4SG7zIyC9kRm6bwyvWCTqb6cr2h-XyI/edit#gid=0';
var RAW_REPORT_TAB_NAME = 'Daily Report';

var CURRENCY_EXCHANGE_TAB_NAME = 'Currency Exchange';

function main() {
  
  var output = [];
  var stats = getStatsByAccount();
  
  var MANAGER_LABELS = getManagerLabels();
  for(var k in MANAGER_LABELS) {
    var iter = MccApp.accounts()
    .withCondition('LabelNames CONTAINS "'+MANAGER_LABELS[k]+'"')
    .withCondition('Cost > 0')
    .forDateRange('YESTERDAY')
    //.forDateRange('20180917','20180917')
    .get();
    
    while(iter.hasNext()){
      MccApp.select(iter.next()) ;
      var id = AdWordsApp.currentAccount().getCustomerId();
      if(!stats[id]) { continue;}
      
      var map = stats[id];
      output.push([map.DATE, map.MONTH, map.QUARTER, MANAGER_LABELS[k], 
      			   id, map.NAME, map.CURRENCY, map.SPENDS, map.GBP_SPENDS]);
      
      delete stats[id];
    }
  }
  
  for(var id in stats) {
    var map = stats[id];
    output.push([map.DATE, map.MONTH, map.QUARTER, 'Other', 
      			 id, map.NAME, map.CURRENCY, map.SPENDS, map.GBP_SPENDS]);
  }
  
  //Logger.log(output);
  //var TEMP_URL = 'https://docs.google.com/spreadsheets/d/13HP0vitNQrc4Dl2QN--CZVvvt0bbASX5PEEsmTKZtZc/edit#gid=223619807';
  var tab = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName(RAW_REPORT_TAB_NAME);
  tab.getRange(tab.getLastRow()+1, 1, output.length, output[0].length).setValues(output);
}


function getStatsByAccount() {
  var quarterMap = { 
    'January': 'Q1', 'February': 'Q1', 'March': 'Q1', 'April': 'Q2', 
    'May': 'Q2', 'June': 'Q2', 'July': 'Q3', 'August': 'Q3', 
    'September': 'Q3', 'October': 'Q4', 'November': 'Q4', 'December': 'Q4' 
  };
  
  
  var mccAccount = AdWordsApp.currentAccount();
  
  var CURRENCY_MAP = readCurrencyExchangeRates();
  var map = readAccountsToIgnore();
  var ids = Object.keys(map);
  
  var dt = '20180917', diff = 1;
  
  var iter = MccApp.accounts()
  .withCondition('Name DOES_NOT_CONTAIN "CSS -"')
  .withCondition('Cost > 0')
  .forDateRange('YESTERDAY')
  //.forDateRange(dt, dt)
  .get();
  
  var statsMap = {};
  while(iter.hasNext()) {
    MccApp.select(iter.next());
    if(!AdWordsApp.currentAccount().getName()) { continue; }
    if(ids.indexOf(AdWordsApp.currentAccount().getCustomerId()) > -1) { continue; }
    
    var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    now.setHours(12);
    now.setDate(now.getDate()-diff);    
    
    var date = Utilities.formatDate(now, 'PST', 'MMM dd, yyyy');
    //var date = 'Jul 27, 2018';
    var month = Utilities.formatDate(now, 'PST', 'MMMM'); 
    
    var quarter = quarterMap[month];
    
    var spends = AdWordsApp.currentAccount()
    .getStatsFor('YESTERDAY')
    //.getStatsFor(dt, dt)
    .getCost();
    
    var currency = AdWordsApp.currentAccount().getCurrencyCode();
    
    var gbpSpends = spends;
    if(currency != 'GBP') {
      gbpSpends = spends*CURRENCY_MAP[currency];
    }
    
    statsMap[AdWordsApp.currentAccount().getCustomerId()] = {
      'NAME': AdWordsApp.currentAccount().getName(), 'CURRENCY': AdWordsApp.currentAccount().getCurrencyCode(),
      'SPENDS': spends, 'GBP_SPENDS': gbpSpends,
      'DATE': date, 'MONTH': month, 'QUARTER': quarter
    };
  }
  
  return statsMap
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

function getManagerLabels() {
  var MANAGER_LABELS = [];
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var URL_INFO_TAB = 'Dashboard Urls';
  
  var urlData = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(URL_INFO_TAB).getDataRange().getValues();
  urlData.shift();
  for(var j in urlData) {
    if(!urlData[j][0] || urlData[j][0] == 'New Business') { continue; }
    MANAGER_LABELS.push(urlData[j][0]);
  }
  
  return MANAGER_LABELS;
}