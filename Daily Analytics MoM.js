var FOLDER_ID = '0B48ewrnCIsYAcEx1ZzFDdGg0c1E',
    TEMPLATE_URL = 'https://docs.google.com/spreadsheets/d/1y6bZg2sNw_WLMKM80urWLmgAQZ8-bAv1DVu0a-c02JE/edit#gid=1155337470',
    TEMPLATE_TAB_NAME = 'Daily Analytics Revenue',
    MASTER_EMAILS = [
      'ricky@pushgroup.co.uk','charlie@pushgroup.co.uk','nj.scripts.mcc@gmail.com',
      'adwords@pushgroup.co.uk','backuppushdomains@gmail.com',
      'analytics@pushgroup.co.uk','master@pushgroup.co.uk','master%pushgroup.co.uk@gtempaccount.com',
      'execbackup@pushdomains.co.uk','charlieppc@pushgroup.co.uk'
    ];

function main() {
  var CONFIG = parseInputs();
  
  for(var id in CONFIG) {
    //if(id != 109239344) { continue; }
    try {
      compileReport(id, CONFIG[id].REPORT_URL, CONFIG[id].CUSTOMER_ID);
    } catch(eee) {
      Logger.log(eee); 
    }
  }
}

function parseInputs() {
  var CONFIG = {};
  var url = 'https://docs.google.com/spreadsheets/d/1wPttqnh4aWkkIRRsfELOBPxKmDf4hTxfgCFRlbnoiro/edit';
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Daily Reports');
  var data = tab.getDataRange().getValues();
  var HEADER = data.shift();
  for(var z in data) {
    var SETTINGS = {};
    for(var i in HEADER) {
      SETTINGS[HEADER[i]] = data[z][i]; 
    }
    
    if(!SETTINGS.REPORT_URL) {
      var ss = SpreadsheetApp.create(SETTINGS.ACCOUNT_NAME + ': Push Performance Report');
      ss.addEditors(MASTER_EMAILS);
      SETTINGS.REPORT_URL = ss.getUrl();
      tab.getRange(parseInt(z,10)+2,3).setValue(SETTINGS.REPORT_URL);
      DriveApp.getFolderById(FOLDER_ID).addFile(DriveApp.getFileById(ss.getId()));
    }
    
    CONFIG[data[z][0]] = SETTINGS;
  }
  
  return CONFIG;
}

function compileReport(id, url, customer_id) {
  var dt = new Date();
  dt.setHours(12);
  
  var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'));
  if(now.getDate() < 3) {
    dt.setDate(0);
    var month = Utilities.formatDate(dt, 'GMT', 'MMM yyyy');
    var end = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    var start = end.substring(0,8) + '01';
    compileReportForMonth(id, start, end, month, url, dt, customer_id); 
  }
  
  var dt = new Date();
  dt.setHours(12);
  var month = Utilities.formatDate(new Date(), 'GMT', 'MMM yyyy');
  var end = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  var start = end.substring(0,8) + '01';
  compileReportForMonth(id, start, end, month, url, dt, customer_id);
}

function compileReportForMonth(id, start, end, month, url, dt, customer_id) {
  var ss = SpreadsheetApp.openByUrl(url);
  var tabName = 'AdWords: ' +month;
  var tab = ss.getSheetByName(tabName);
  if(!tab) {
    var tab;
    
    if(id != '104503053') {
      var dummy = SpreadsheetApp.openByUrl(TEMPLATE_URL).getSheetByName(TEMPLATE_TAB_NAME);
      tab = dummy.copyTo(ss);
    } else {
      ss.setActiveSheet(ss.getSheetByName('Template'));
      tab = ss.duplicateActiveSheet();
    }
    
    tab.setName(tabName);
    tab.showSheet();
  }
  
  var initMap = {
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0,
    'AssistedConversions': 0, 'AssistedConversionValue': 0, 'Sessions': 0
  };
  
  var stats = {};
  var sd = new Date(start),
      ed = new Date(end);
  while(sd <= ed) {
    stats[Utilities.formatDate(sd, 'GMT', 'yyyyMMdd')] = JSON.parse(JSON.stringify(initMap));
    sd.setDate(sd.getDate()+1);
  }
  
  var filters = ['ga:medium==cpc;ga:source==google']
  if(customer_id) {
    filters.push('ga:adwordsCustomerID==' + customer_id); 
  }
  
  var optArgs = { 
    'dimensions': 'ga:date', 
    'filters': filters.join(';')
  };
  
  getDataFromAnalytics(id,stats,initMap,start,end,optArgs, customer_id);
  getDataFromMCF(id,stats,initMap,start,end, customer_id);
  
  var col = 3;
  for(var date in stats) {
    col++;
    var rows = [[stats[date].Cost], [stats[date].ConversionValue], [stats[date].Conversions],
                [stats[date].AssistedConversions], [stats[date].AssistedConversionValue]];
    tab.getRange(4, col, rows.length, 1).setValues(rows);
  }
}

function getDataFromAnalytics(PROFILE_ID,stats,initMap,FROM,TO,optArgs,customer_id) {
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue,ga:sessions",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    if(!stats[rows[k][0]]) {
      stats[rows[k][0]] = JSON.parse(JSON.stringify(initMap));
    }
    
    stats[rows[k][0]].Cost += parseFloat(rows[k][1]);
    stats[rows[k][0]].Conversions += parseInt(rows[k][2],10);
    stats[rows[k][0]].ConversionValue += parseFloat(rows[k][3]);
    stats[rows[k][0]].Sessions += parseInt(rows[k][4],10);    
  }
}

function getDataFromMCF(PROFILE_ID,stats,initMap,FROM, TO,customer_id) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  if(customer_id) {
    filters.push('mcf:adwordsCustomerID==' + customer_id); 
  }
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType,mcf:conversionDate',
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                  // End-date (format yyyy-MM-dd).
    "mcf:totalConversions,mcf:totalConversionValue",
    optArgs
  );
  
  var rows = results.rows;
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    var dt = rows[k][2].primitiveValue;
    if(!stats[dt]) {
      stats[dt] = JSON.parse(JSON.stringify(initMap));
    }
    
    stats[dt].AssistedConversions += parseInt(rows[k][3].primitiveValue,10);
    stats[dt].AssistedConversionValue += parseFloat(rows[k][4].primitiveValue);    
  }
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}