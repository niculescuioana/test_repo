var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var TAB_NAME = 'Dashboard Urls';

var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1TVNX_V4SA928eTqqDafqGNSaA9cgOzeeedmAtKBH9xg/edit';

function main() {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  
  var diffFlag = 0;
  var today = getAdWordsFormattedDate(-1+diffFlag, 'd MMMM');
  var date = ordinal_suffix_of(parseInt(today.split(' ')[0],10));
  var tabName = date + ' ' + today.split(' ')[1];
 
  var accountMap = readAllAccounts();
  
  var ss = SpreadsheetApp.openByUrl(REPORT_URL);
  var tab = ss.getSheetByName(tabName);
  
  if(!tab) {
    var lastWeek = getAdWordsFormattedDate(6+diffFlag, 'd MMMM');
    var date = ordinal_suffix_of(parseInt(lastWeek.split(' ')[0],10));
    var lastWeekTabName = date + ' ' + lastWeek.split(' ')[1];
    var lwTab = ss.getSheetByName(lastWeekTabName);
    
    ss.setActiveSheet(lwTab);
    tab = ss.duplicateActiveSheet();
    tab.setName(tabName);
    ss.setActiveSheet(tab);
    ss.moveActiveSheet(1);
    
    tab.getRange(1,1).setValue(tabName);
    
    var accountsByManagers = getAccountsByManagers();
    
    var data = tab.getDataRange().getValues();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    
    
    var manager = '', managerStartRow = {}, toDelete = {};
    for(var k in data) {
      if(accountsByManagers[data[k][0]]) {
        manager = data[k][0]; 
        managerStartRow[manager] = parseInt(k,10)+9; 
      }
      
      if(!manager || !data[k][0] || !accountsByManagers[manager] || accountsByManagers[data[k][0]] 
         || data[k][0] == 'General Task' || data[k][0] == 'Account')  { continue; }
      
      if(!accountsByManagers[manager][data[k][0]]) {
        if(!toDelete[manager]) {
          toDelete[manager] = {};
        }
        
        toDelete[manager][data[k][0]] = 1;
        continue;
      }
      
      delete accountsByManagers[manager][data[k][0]];
    }
    
    var keys = Object.keys(managerStartRow);
    
    for(var i = keys.length-1; i >= 0; i--) {
      var manager = keys[i];
      if(!accountsByManagers[manager]) { continue; }
      
      for(var account in accountsByManagers[manager]) {
        tab.insertRowAfter(managerStartRow[manager]);
        tab.getRange(managerStartRow[manager]+1, 1).setValue(account);
      }
    }
    
    var data = tab.getDataRange().getValues();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    
    var manager = '', delRows = [];
    for(var k in data) {
      if(toDelete[data[k][0]]) {
        manager = data[k][0];
        continue;
      }
      
      if(toDelete[manager] && toDelete[manager][data[k][0]]) {
        delRows.push(parseInt(k,10) + 7); 
      }
    }
    
    for(var l=delRows.length-1; l>= 0; l--) {
      tab.deleteRow(delRows[l]);
    }
    
    //Logger.log(delRows);
  } 
  
  var backgrounds = [];
  var whiteBgRow = ['#ffffff', '#ffffff', '#ffffff'];
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  data.shift();
  
  for(var k in data) {
    if(!data[k][0] || !accountMap[data[k][0]]) { 
      backgrounds.push(whiteBgRow);
      continue; 
    }
    
    var bgRow = ['#ffffff', '#ffffff', '#ffffff'];
    if(!data[k][2]) {
      bgRow[0] = '#dd7e6b';
    }
    
    if(!data[k][3]) {
      bgRow[1] = '#dd7e6b';
    }
    
    if(!data[k][3]) {
      bgRow[2] = '#dd7e6b';
    }
    
    backgrounds.push(bgRow);
  }
  
  //Logger.log(backgrounds);
  
  tab.getRange(7,3,backgrounds.length,backgrounds[0].length).setBackgrounds(backgrounds);
  
  
  tab.getRange('C7:E325').clearContent();
  tab.getRange('J7:N325').clearContent();
  tab.getRange('A2:A7').clearContent();
  
  cleanTabs(ss);
}

function cleanTabs(ss) {
  var month_yest =  getAdWordsFormattedDate(1, 'MMMM');
  var month_today =  getAdWordsFormattedDate(-1, 'MMMM');  
  var last_month =  getAdWordsFormattedDate(31, 'MMMM');
  
  
  var tabs = ss.getSheets();
  for(var z in tabs) {
    var name = tabs[z].getName();
    if(name.indexOf(month_yest) > -1 || name.indexOf(month_today) > -1 || name.indexOf(last_month) > -1) {
      // ok
    } else {
      ss.deleteSheet(tabs[z]) ;
    }
  }
  
}

function getAccountsByManagers() {
  var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(TAB_NAME);
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var accountsByManagers = {};
  for(var k in data) {
    var managerLabel = data[k][0];
    var manager = data[k][0];
    accountsByManagers[manager] = {};
    var accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "' + managerLabel + '"').get();
    while(accounts.hasNext()) {
      var account = accounts.next();
      accountsByManagers[manager][account.getName()] = 1;
    }
  } 
  
  return accountsByManagers;
}

function readAllAccounts() {
  var accountMap = {};
  var accounts = MccApp.accounts().get();
  while(accounts.hasNext()) {
    var account = accounts.next();
    accountMap[account.getName()] = 1;
  }
  
  return accountMap;
}

function ordinal_suffix_of(i) {
  var j = (i % 10), 
      k = (i % 100);
  
  if (j == 1 && k != 11) {
    return i + "st";
  }
  if (j == 2 && k != 12) {
    return i + "nd";
  }
  if (j == 3 && k != 13) {
    return i + "rd";
  }
  return i + "th";
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function info(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}
