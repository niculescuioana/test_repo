var URL = 'https://docs.google.com/spreadsheets/d/1N73pAX3ARTGhlueYlF2F24wd-_eyK7AH1fAC7yFGRY8/edit#gid=1978516087';

function main() {
  var NEW_SPREADSHEET_ROWS = [
    ['Duplicate Keyword Report', '', 'DUPLICATE_KEYWORDS_EXPORT_URL', 'TRUE'],
    ['Excessive Ads Report', '', 'EXTRA_ADS_ADGROUPS_REPORT_URL', 'TRUE'],
    ['Negative Upload Spreadsheet', '', 'NEGATIVE_UPLOAD_SS_URL', 'TRUE'],
    ['Spell Check Report', '', 'SPELL_CHECK_URL', 'TRUE'],
    ['Trends Overview', '', 'TRENDS_OVERVIEW_URL', 'TRUE'],
    ['Weekly Benchmark Report', '', 'WEEKLY_BENCHMARK_REPORT_URL', 'TRUE'],
    ['Holiday Manager Report', '', 'HOLIDAY_MANAGER_URL', 'FALSE'],
    ['Underperforming Placements Report', '', 'UNDERPERFORMING_PLACEMENTS_URL', 'TRUE'],
    ['Negatives Clashing Keywords', '', 'NEGATIVES_CLASHING_KEYWORDS_URL', 'TRUE'],
    ['Quarterly Business Review', '', 'MOM_METRIC_COMPARISON_URL', 'TRUE'],
    ['Ad Schedule Bidding Report', '', 'AD_SCHEDULE_REPORT_URL', 'TRUE']
  ];
  
  for(var z in NEW_SPREADSHEET_ROWS) {
    addRows(NEW_SPREADSHEET_ROWS[z]);
  }
}




function addRows(NEW_SPREADSHEET_ROW) {
  var MANAGER = 'Elliot (Navico)';
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Report Links');
  var input = tab.getDataRange().getValues();
  
  var found = false;
  for(var x in input) {
    if(NEW_SPREADSHEET_ROW[0] == input[x][0]) {
      found = true;
    } 
  }
  
  if(!found) {
    var ss = SpreadsheetApp.create(NEW_SPREADSHEET_ROW[0] + ' - ' + MANAGER);
    NEW_SPREADSHEET_ROW[1] = ss.getUrl();
    tab.appendRow(NEW_SPREADSHEET_ROW);
  }
}

function addRowsToAll(NEW_SPREADSHEET_ROW) {
  var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
  var DASHBOARD_TAB_NAME = 'Dashboard Urls';
  var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    var MANAGER = data[k][0];
    var URL = data[k][1];
    var ss = SpreadsheetApp.openByUrl(URL); 
    
    var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Report Links');
    var input = tab.getDataRange().getValues();
    
    var found = false;
    for(var x in input) {
      if(NEW_SPREADSHEET_ROW[0] == input[x][0]) {
        found = true;
      } 
    }
    
    if(!found) {
      var ss = SpreadsheetApp.create(NEW_SPREADSHEET_ROW[0] + ' - ' + MANAGER);
      NEW_SPREADSHEET_ROW[1] = ss.getUrl();
      tab.appendRow(NEW_SPREADSHEET_ROW);
    }
  }
  
}

function temp() {
  //MccApp.accounts().withCondition('Name = "Barriers Direct"').executeInParallel('runBarriers2');
  
  var map = {};
  var iter = MccApp.accounts().get();
  while(iter.hasNext()) {
    MccApp.select(iter.next());
    map[AdWordsApp.currentAccount().getCustomerId().replace(/-/g, '')] = AdWordsApp.currentAccount().getName();
  }
  
  var url = 'https://docs.google.com/spreadsheets/d/14M-2XKWLYiH_n7BsdgyzHuGmd_N7YPRvWHhwclc3CWs/edit#gid=1995295753';
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Accounting CIDs');
  var data = tab.getRange(2,1,897,3).getValues();
  for(var z in data) {
    if(!map[data[z][2]]) { continue;}
    data[z][0] = map[data[z][2]];
  }
  
  tab.getRange(2,1,897,3).setValues(data);
}

/*function runBarriers1() {
  var map = {};
  
  var query = [
    'SELECT CampaignName, Criteria FROM KEYWORDS_PERFORMANCE_REPORT',
    'WHERE Status = ENABLED and AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
    'DURING YESTERDAY'
  ].join(' ');
  
  var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    if(!map[row.CampaignName]) {
      map[row.CampaignName] = {}; 
    }
    
    map[row.CampaignName][row.Criteria.toLowerCase().replace(/[+]/g, '')] = 1;
  }
  
  var URL = 'https://docs.google.com/spreadsheets/d/1cvPTMTX8LUJi5jLMU4yRPvnBPgtadbhpyEAhLlhWLX4/edit#gid=0';
  var tab = SpreadsheetApp.openByUrl(URL).getSheets()[0];
  var col = 1;
  for(var camp in map) {
    var out = [[camp]];
    for(var kw in map[camp]) {
      out.push([kw]) ;
    }
    
    tab.getRange(1,col,out.length,1).setValues(out);
    col++;
  }
}

function runBarriers2() {
  var URL = 'https://docs.google.com/spreadsheets/d/1cvPTMTX8LUJi5jLMU4yRPvnBPgtadbhpyEAhLlhWLX4/edit#gid=883974207';
  var map = {};
  
  var query = [
    'SELECT CampaignName, AdGroupName, AdGroupId, Criteria, KeywordMatchType FROM KEYWORDS_PERFORMANCE_REPORT',
    'WHERE Status = ENABLED and AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
    'DURING YESTERDAY'
  ].join(' ');
  
  var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    if(!map[row.AdGroupId]) {
      map[row.AdGroupId] = {
        'Campaign': row.CampaignName,
        'AdGroup': row.AdGroupName,
        'Keywords': {}
      }; 
    }
    
    row.Keyword = row.Criteria.toLowerCase().replace(/[+]/g, '');
    if(!map[row.AdGroupId]['Keywords'][row.Keyword]) {
      map[row.AdGroupId]['Keywords'][row.Keyword] = {
        'Exact': 0, 'Phrase': 0, 'Broad': 0 
      }
    }
    
    map[row.AdGroupId]['Keywords'][row.Keyword][row.KeywordMatchType] = 1;
  }
  
  var out = [];
  for(var id in map) {
    for(var kw in map[id]['Keywords']) {
      var row = map[id]['Keywords'][kw];
      if(!row.Exact || !row.Phrase || !row.Broad) {
        out.push([map[id]['Campaign'], map[id]['AdGroup'], kw, row.Exact ? 'Y' : 'N', row.Phrase ? 'Y' : 'N', row.Broad ? 'Y' : 'N']); 
      }
    }
  }
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheets()[1];
  tab.getRange(2,1,out.length,out[0].length).setValues(out);
}
*/

/*
var NEW_SCRIPT_ROWS = [
['Ad Schedule Manager', 'AdScheduleManager_v2.js', 'runAdScheduleManager', '', 'Ad Schedule Manager v2', 'Hourly', '', '', '']
];

var TAB_NAME = 'Scripts Config 2';

var NEW_SPREADSHEET_ROWS = ['Seasonality Checker', 'https://docs.google.com/spreadsheets/d/1FRoaKJ3WFZ0mDVJObDflwd6K3NrG0i_EQ63cS1Rx4to/edit#gid=0', '', 'FALSE'];




function main() {
//temp();
//step1();

copyStep();
}

function copyStep() {
var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_TAB_NAME = 'Dashboard Urls';
var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME).getDataRange().getValues();
data.shift();

for(var k in data) {
var MANAGER = data[k][0];
var URL = data[k][1];
var ss = SpreadsheetApp.openByUrl(URL);
if(ss.getSheetByName('QS Tracker Report')) { continue; }
if(!ss.getSheetByName('Quality Score Report')) {
Logger.log(MANAGER);
continue;
}
ss.setActiveSheet(ss.getSheetByName('Quality Score Report'));
var tab = ss.duplicateActiveSheet();
tab.setName('QS Tracker Report');
tab.getRange('B1').setValue('QS TRACKER')
tab.getRange('B3:C').clearContent();
}
}

function step2() {
var name = 'Seasonality Check';
var monique = 'https://docs.google.com/spreadsheets/d/1ldQ234BVUh2WleIVDRvsgbrJafLuP7fq7weDvKFbj0s/edit#gid=169635450';
var tab = SpreadsheetApp.openByUrl(monique).getSheetByName(name);
var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_TAB_NAME = 'Dashboard Urls';
var data = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME).getDataRange().getValues();
data.shift();

for(var k in data) {
var MANAGER = data[k][0];
if(MANAGER == 'Isuru') { continue; }
var URL = data[k][1];
var ss = SpreadsheetApp.openByUrl(URL);
var dummy = ss.getSheetByName(name)
if(dummy) {
ss.deleteSheet(dummy);
}

continue;

var newTab = tab.copyTo(ss);
newTab.setName(name);

newTab.getRange(3,1,newTab.getLastRow(),newTab.getLastColumn()).clearContent();
SpreadsheetApp.openByUrl(URL).getSheetByName('Management Report').getRange('A:A').copyTo(newTab.getRange('A:A'));
}
}

function step1() {
var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_TAB_NAME = 'Dashboard Urls';
var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME) ;
var data = tab.getDataRange().getValues();
data.shift();

for(var k in data) {
var MANAGER = data[k][0];
var URL = data[k][1];
Logger.log(MANAGER + ' : '  +URL);

var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Report Links');
tab.appendRow(NEW_SPREADSHEET_ROWS);

//copyConfigRow(URL, MANAGER);
//copyGeneralInputRow(URL, MANAGER);
} 
}

function copyConfigRow(URL, MANAGER) {
var tab = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
var data = tab.getDataRange().getValues();

var found = {};
for(var x in data) {
for(var y in NEW_SCRIPT_ROWS) {
if(NEW_SCRIPT_ROWS[y][0] == data[x][0]) {
found[NEW_SCRIPT_ROWS[y][0]] = 1;
}
}
}

for(var y in NEW_SCRIPT_ROWS) {
if(!found[NEW_SCRIPT_ROWS[y][0]]) {
tab.appendRow(NEW_SCRIPT_ROWS[y]);
}
}
}

function copyGeneralInputRow(URL, MANAGER) {
var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('General Inputs');
var data = tab.getDataRange().getValues();

var found = {};
for(var x in data) {
for(var y in NEW_SPREADSHEET_ROWS) {
if(NEW_SPREADSHEET_ROWS[y][0] == data[x][0] && data[x][1]) {
found[NEW_SPREADSHEET_ROWS[y][0]] = 1;
}
}
}

for(var y in NEW_SPREADSHEET_ROWS) {
if(!found[NEW_SPREADSHEET_ROWS[y][0]]) {
var ss = SpreadsheetApp.create(NEW_SPREADSHEET_ROWS[y][0] + ' - ' + MANAGER);
NEW_SPREADSHEET_ROWS[y][1] = ss.getUrl();

tab.appendRow(NEW_SPREADSHEET_ROWS[y]);
}
}
}



function temp() {
var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_TAB_NAME = 'Dashboard Urls';
var tab = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(DASHBOARD_TAB_NAME) ;
var data = tab.getDataRange().getValues();
data.shift();

for(var k in data) {
var MANAGER = data[k][0];
var URL = data[k][1];

var ss = SpreadsheetApp.openByUrl(URL);
var tab = ss.getSheetByName('Daily Snapshot');
if(!tab) { continue; }

ss.deleteSheet(tab);
}
}
*/