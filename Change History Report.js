var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var TAB_NAME = 'Dashboard Urls';

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1LvPXAfeTM8Fss3rFVGUhlUth8pWM6HFxpnc0SRdU1T4/edit';
var TEMPLATE_TAB_NAME = 'Template';

function main() {
  
  
  var tabName = getDateRange();
  var ss = SpreadsheetApp.openByUrl(REPORT_URL);
  if(ss.getSheetByName(tabName)) { return; }
  
  var accountTierMap = {};
  var tierLabels = ['Tier 1', 'Tier 2'];
  
  for(var k in tierLabels) {
    var label = tierLabels[k];
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "' + label + '"').get();
    while(accounts.hasNext()) {
      var account = accounts.next();
      accountTierMap[account.getName()] = { 'Tier': label, 'Managers': [] };
    }
  }
  
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME) ;
  var data = tab.getDataRange().getValues();
  data.shift();
  
  
  for(var k in data) {
    var manager = data[k][0]
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "' + manager + '"').get();
    while(accounts.hasNext()) {
      var account = accounts.next();
      if(!accountTierMap[account.getName()]) { continue; }
      accountTierMap[account.getName()]['Managers'].push(manager);
    }
  }
  
  var output = [];
  for(var accountName in accountTierMap) {
    output.push([accountTierMap[accountName]['Tier'], accountName, accountTierMap[accountName]['Managers'].join('/')]);
  }
  
  var templateTab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Template');
  ss.setActiveSheet(templateTab);
  var tab = ss.duplicateActiveSheet();
  tab.setName(tabName);
  tab.showSheet();
  
  tab.getRange(4,1,output.length,output[0].length).setValues(output).sort([{'column': 1, 'ascending': true}, {'column': 3, 'ascending': true}])
  tab.getRange(1,2).setValue(tabName);
}

function getDateRange() {
  var start = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var diff = start.getDay()-1;
  start.setDate(start.getDate()-diff);
  
  var end = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var diff = end.getDay()-1;
  end.setDate(end.getDate()-diff);
  end.setDate(end.getDate()+4);
  
  var dateRange;
  if(end.getMonth() == start.getMonth()) {
    dateRange = Utilities.formatDate(start, 'PST', 'd') + '-' + Utilities.formatDate(end, 'PST', 'd MMM yyyy');
  } else {
    dateRange = Utilities.formatDate(start, 'PST', 'd MMM') + '-' + Utilities.formatDate(end, 'PST', 'd MMM yyyy'); 
  }
  
  return dateRange;
}