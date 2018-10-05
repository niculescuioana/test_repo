var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1elAfLK9py2HCI07I1FsKkzq5gv8j0I5UnehWqcRAbww/edit?ts=5a0eb34a#gid=0';

function main() {

  var YESTERDAY = getAdWordsFormattedDate_(1, 'yyyy-MM-dd');
  var START = YESTERDAY.substring(0,8) + '01';
  
  var parts = YESTERDAY.split('-');
  var ly = parseInt(parts.shift(), 10) - 1;
  
  var date = new Date(getAdWordsFormattedDate_(1, 'MMM d, yyyy'));
  date.setDate(1);
  
  date.setDate(0);
  var LMEND = Utilities.formatDate(date, 'GMT', 'yyyy-MM-dd');
  
  date.setDate(1);
  var LMSTART = Utilities.formatDate(date, 'GMT', 'yyyy-MM-dd');
  
  var date = new Date(getAdWordsFormattedDate_(1, 'MMM d, yyyy'));
  date.setYear(date.getYear()-1);
  date.setDate(1);
  date.setMonth(date.getMonth()+1);
  date.setDate(0);
  date.setHours(12);
  
  var LYEND = Utilities.formatDate(date, 'GMT', 'yyyy-MM-dd');
  
  var DATES = {
    'YESTERDAY': { 'START': getAdWordsFormattedDate_(1, 'yyyy-MM-dd'), 'END': getAdWordsFormattedDate_(1, 'yyyy-MM-dd') },
    'TODAY': { 'START': getAdWordsFormattedDate_(0, 'yyyy-MM-dd'), 'END': getAdWordsFormattedDate_(0, 'yyyy-MM-dd') },
    'LM': { 'START': LMSTART, 'END': LMEND },
    'MTD': { 'START': START, 'END': YESTERDAY },
    'LMTD': { 'START': ly+'-'+parts[0]+'-01', 'END': ly+'-'+parts.join('-') },
    'LYM': {  'START': ly+'-'+parts[0]+'-01', 'END': LYEND }
  };
  
  //Logger.log(DATES);
  //return;
  var map = {};
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Report');
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  
  for(var z in data) {
    if(!data[z][0]) { continue; }
    runForProfile_(data[z][0], parseInt(z,10)+4, DATES, map);
  }
  
}

function runForProfile_(PROFILE_ID, ROW_NUM, DATES, map) {
  var initMap = {};
  for(var key in DATES) {
    initMap[key] = { 'Cost': 0, 'ConversionValue': 0, 'Transactions': 0 }; 
  }
  
  var stats = map[PROFILE_ID];
  if(!stats) {
    stats = {
      'ALL': JSON.parse(JSON.stringify(initMap)),
      'ADWORDS': JSON.parse(JSON.stringify(initMap))
    };
    
    var optArgs;
    for(var key in DATES) {
      optArgs = {};
      getDataFromAnalytics_(PROFILE_ID, stats['ALL'][key], DATES[key].START, DATES[key].END, optArgs);
      
      if(stats['ALL'][key]['ConversionValue'] == 0) { continue; }
      
      optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
      getDataFromAnalytics_(PROFILE_ID, stats['ADWORDS'][key], DATES[key].START, DATES[key].END, optArgs);
      
      getDataFromMCF_(PROFILE_ID, stats['ADWORDS'][key], DATES[key].START, DATES[key].END);
    }
  }
  
  var rows = [];
  for(var key in stats) {
    var row = [];
    for(var date in stats[key]) {
      var cpt = stats[key][date].Transactions == 0 ? 0 : round_(stats[key][date].Cost / stats[key][date].Transactions, 2);
      var roas = stats[key][date].Cost == 0 ? 0 : round_(stats[key][date].ConversionValue / stats[key][date].Cost, 4);
      row.push(stats[key][date].ConversionValue, stats[key][date].Transactions, stats[key][date].Cost, cpt, roas);
    }
    rows.push(row);
  }
  
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Report');
  tab.getRange(ROW_NUM, 5, rows.length, rows[0].length).setValues(rows);              
  
  SpreadsheetApp.flush();
}

function getDataFromAnalytics_(PROFILE_ID,stats,FROM,TO,optArgs) {
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactionRevenue,ga:transactions",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex);
      if(ex.message.indexOf('permissions') > -1) { attempts = 0; break; }
      attempts--;
      Utilities.sleep(2500);
    }
  }
  
  if(!attempts) { return; }
  
  var rows = resp.getRows();
  for(var k in rows) {
    stats.Cost += parseFloat(rows[k][0]);
    stats.ConversionValue += parseFloat(rows[k][1]);
    stats.Transactions += parseInt(rows[k][2], 10);
  }
}


function getDataFromMCF_(PROFILE_ID, stats, FROM, TO) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType',
    'filters': filters.join(';')
  };
  
  try {
    var results = Analytics.Data.Mcf.get(
      'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
      FROM,                 // Start-date (format yyyy-MM-dd).
      TO,                  // End-date (format yyyy-MM-dd).
      "mcf:totalConversionValue,mcf:totalConversions",
      optArgs
    );
  } catch(ex) {
    Logger.log(ex);
    return;
  }
  
  var rows = results.rows;
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    //stats.AssistedConversions += parseInt(rows[k][2].primitiveValue,10);
    stats.ConversionValue += parseFloat(rows[k][2].primitiveValue);    
    stats.Transactions += parseInt(rows[k][3].primitiveValue, 10);    
  }
}

function round_(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}

function getAdWordsFormattedDate_(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,'GMT',format);
}
