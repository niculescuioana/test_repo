var FOLDER_ID = '0B48ewrnCIsYAcEx1ZzFDdGg0c1E',
    TEMPLATE_URL = 'https://docs.google.com/spreadsheets/d/1y6bZg2sNw_WLMKM80urWLmgAQZ8-bAv1DVu0a-c02JE/edit#gid=1155337470',
    TEMPLATE_TAB_NAME = 'Monthly Analytics Revenue',
    MASTER_EMAILS = [
      'ricky@pushgroup.co.uk','charlie@pushgroup.co.uk','nj.scripts.mcc@gmail.com',
      'adwords@pushgroup.co.uk','backuppushdomains@gmail.com',
      'analytics@pushgroup.co.uk','master@pushgroup.co.uk','master%pushgroup.co.uk@gtempaccount.com',
      'execbackup@pushdomains.co.uk','charlieppc@pushgroup.co.uk'
    ];

function main() {
  var CONFIG = parseInputs();
  MccApp.accounts()
  .withIds(Object.keys(CONFIG))
  //.withIds(['855-699-6922'])
  .executeInParallel('run', 'compile', JSON.stringify(CONFIG));
}

function compile() {}

function run(input) {
  
  var INPUTS = JSON.parse(input)[AdWordsApp.currentAccount().getCustomerId()];
  for(var i in INPUTS) {
    runForSetting(INPUTS[i]);
  }
}

function parseInputs() {
  var CONFIG = {}, URL_MAP = {};
  var url = 'https://docs.google.com/spreadsheets/d/1wPttqnh4aWkkIRRsfELOBPxKmDf4hTxfgCFRlbnoiro/edit';
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Monthly Reports');
  var data = tab.getDataRange().getValues();
  var HEADER = data.shift();
  for(var z in data) {
    if(!CONFIG[data[z][0]]) {
      CONFIG[data[z][0]] = [];
    }
    
    var SETTINGS = {};
    for(var i in HEADER) {
      SETTINGS[HEADER[i]] = data[z][i]; 
    }
    
    if(!SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN']) {
      SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN'] = [] 
    } else {
      SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN'] = SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN'].split(',');
      for(var k in SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN']) {
        SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN'][k] = SETTINGS['CAMPAIGN_DOES_NOT_CONTAIN'][k].trim();
      }
    }
    
    SETTINGS.FILTER = '';
    
    if(!SETTINGS.REPORT_URL) {
      if(URL_MAP[SETTINGS.ACCOUNT_ID]) {
        SETTINGS.REPORT_URL = URL_MAP[SETTINGS.ACCOUNT_ID]; 
      } else {
        var ss = SpreadsheetApp.create(SETTINGS.ACCOUNT_NAME + ' YoY');
        ss.addEditors(MASTER_EMAILS);
        SETTINGS.REPORT_URL = ss.getUrl();
        DriveApp.getFolderById(FOLDER_ID).addFile(DriveApp.getFileById(ss.getId()));
      }
      
      tab.getRange(parseInt(z,10)+2,5).setValue(SETTINGS.REPORT_URL);
    }
    
    URL_MAP[SETTINGS.ACCOUNT_ID] = SETTINGS.URL;
    
    CONFIG[data[z][0]].push(SETTINGS);
  }
  
  return CONFIG;
}

function runForSetting(SETTINGS) {
  //Logger.log(SETTINGS);
  compileReport(SETTINGS.PROFILE, SETTINGS.REPORT_URL, SETTINGS.CAMPAIGN_CONTAINS, SETTINGS.CAMPAIGN_DOES_NOT_CONTAIN, SETTINGS.LABEL, SETTINGS.FILTER, SETTINGS.TAB_NAME, SETTINGS.OVERALL_REVENUE_FLAG, SETTINGS.CUSTOMER_ID);
  
  if(SETTINGS.WEEKLY_MONTHLY_REPORT_URL) {
    compileWeeklyMonthlyReports(SETTINGS.PROFILE, SETTINGS.WEEKLY_MONTHLY_REPORT_URL, SETTINGS.OVERALL_REVENUE_FLAG);
  }
  
  compileYoYReport(SETTINGS.PROFILE, SETTINGS.REPORT_URL, SETTINGS.CAMPAIGN_CONTAINS, SETTINGS.CAMPAIGN_DOES_NOT_CONTAIN, SETTINGS.LABEL, SETTINGS.FILTER, SETTINGS.OVERALL_REVENUE_FLAG);
  
  if(AdWordsApp.currentAccount().getCustomerId() == '824-654-6369') {
    addConversionByNameStats(SETTINGS.REPORT_URL);
  }
}

function compileYoYReport(id, url, contains, doesNotContain, labelName, filter, overall_flag) {  
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('YoY');
  if(!tab) { return; }
  
  var MONTH = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM');
  var end = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  
  var year = parseInt(end.substring(0,4), 10);
  var sm = end.substring(0,8) + '01';
  
  var sy = year +'-01-01';
  
  var ly = year-1;
  var endLy = ly + end.substring(4,10);
  
  var slym = endLy.substring(0,8) + '01';
  var slyy = ly +'-01-01';
  
  
  var campaignMap = {};
  if(labelName) {
    var iter = AdWordsApp.campaigns().withCondition('LabelNames CONTAINS_ANY ["' + labelName + '"]').get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
    
    var iter = AdWordsApp.shoppingCampaigns().withCondition('LabelNames CONTAINS_ANY ["' + labelName + '"]').get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
  }
  
  if(doesNotContain.length) {
    var iter = AdWordsApp.campaigns();
    for(var z in doesNotContain) {
      iter.withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + doesNotContain[z] + '"');
    }
    
    iter = iter.get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
    
    var iter = AdWordsApp.shoppingCampaigns();
    for(var z in doesNotContain) {
      iter.withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + doesNotContain[z] + '"');
    }
    
    iter = iter.get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
  }
  
  
  var initMap = {
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0,
    'AssistedConversions': 0, 'AssistedConversionValue': 0, 'Sessions': 0,
    'Clicks': 0, 'NewUsers': 0, 'OverallRevenue': 0, 'Bounces': 0
  };
  
  var stats = {
    'ytd': JSON.parse(JSON.stringify(initMap)) ,
    'lytd': JSON.parse(JSON.stringify(initMap)) ,
    'mtd': JSON.parse(JSON.stringify(initMap)) ,
    'lymtd': JSON.parse(JSON.stringify(initMap))
  }
  
  var dates = {
    'ytd': { 'start': sy, 'end': end },
    'lytd': { 'start': slyy, 'end': endLy },
    'mtd': { 'start': sm, 'end': end },
    'lymtd': { 'start': slym, 'end': endLy }
  }
  
  var filters = ['ga:medium==cpc;ga:source==google']
  if(contains) {
    filters.push('ga:campaign=@' + contains);
  }
  
  if(filter) {
    filters.push(filter);
  }
  
  for(var key in stats) {
    if(!Object.keys(campaignMap).length) {
      var optArgs = { 'filters': filters.join(';') };
      getDataFromAnalytics(id,stats[key],dates[key].start,dates[key].end,optArgs,overall_flag);
      getDataFromMCF(id,stats[key],dates[key].start,dates[key].end);
    } else {
      var optArgs = { 'dimensions': 'ga:campaign', 'filters': filters.join(';') };
      getCampaignDataFromAnalytics(id,stats[key],dates[key].start,dates[key].end,optArgs,campaignMap);
      getCampaignDataFromMCF(id,stats[key],dates[key].start,dates[key].end,campaignMap);
    }
  }
  
  for(var key in stats) {
    stats[key]['BR'] = stats[key]['Sessions'] == 0 ? 0 : stats[key]['Bounces'] / stats[key]['Sessions'];
    stats[key]['CPC'] = stats[key]['Clicks'] == 0 ? 0 : stats[key]['Cost'] / stats[key]['Clicks'];
    stats[key]['CR'] = stats[key]['Clicks'] == 0 ? 0 : stats[key]['Conversions'] / stats[key]['Clicks'];
  }
  
  var out = [
    [ly, year],
    [stats['lytd']['Clicks'], stats['ytd']['Clicks']],
    [stats['lytd']['Cost'], stats['ytd']['Cost']],
    [stats['lytd']['CPC'], stats['ytd']['CPC']],
    [stats['lytd']['BR'], stats['ytd']['BR']],
    [stats['lytd']['CR'], stats['ytd']['CR']],
    [stats['lytd']['Conversions'], stats['ytd']['Conversions']],
    ['', ''],
    [stats['lytd']['ConversionValue'], stats['ytd']['ConversionValue']],
    [stats['lytd']['AssistedConversionValue'], stats['ytd']['AssistedConversionValue']]
  ];
  
  tab.getRange(2,2,out.length,out[0].length).setValues(out);
  
  
  var out = [
    [MONTH + ' ' + ly, MONTH + ' ' + year],
    [stats['lymtd']['Clicks'], stats['mtd']['Clicks']],
    [stats['lymtd']['Cost'], stats['mtd']['Cost']],
    [stats['lymtd']['CPC'], stats['mtd']['CPC']],
    [stats['lymtd']['BR'], stats['mtd']['BR']],
    [stats['lymtd']['CR'], stats['mtd']['CR']],
    [stats['lymtd']['Conversions'], stats['mtd']['Conversions']],
    ['', ''],
    [stats['lymtd']['ConversionValue'], stats['mtd']['ConversionValue']],
    [stats['lymtd']['AssistedConversionValue'], stats['mtd']['AssistedConversionValue']]
  ];
  
  tab.getRange(2,6,out.length,out[0].length).setValues(out);
  
  tab.getRange(2,1).setValue(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
}

function addConversionByNameStats(url) {
  var dt = new Date();
  dt.setHours(12);
  
  var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'));
  if(now.getDate() < 3) {
    dt.setDate(0);
    var month = Utilities.formatDate(dt, 'GMT', 'MMM');
    var end = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    var start = end.substring(0,8) + '01';
    addConversionByNameStatsForMonth(start, end, month, url, dt); 
  }
  
  var dt = new Date();
  dt.setHours(12);
  var month = Utilities.formatDate(new Date(), 'GMT', 'MMM');
  var end = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  var start = end.substring(0,8) + '01';
  addConversionByNameStatsForMonth(start, end, month, url, dt);
}

function addConversionByNameStatsForMonth(start, end, month, url, dt) {
  var ty = Utilities.formatDate(dt, 'GMT', 'yyyy');
  var ss = SpreadsheetApp.openByUrl(url);
  
  var tabName = ty;
  var tab = ss.getSheetByName(tabName);
  if(!tab) {
    return;
  }
  
  var stats = {};
  var query = 'SELECT ConversionTypeName, Conversions FROM ACCOUNT_PERFORMANCE_REPORT DURING ' + start.replace(/-/g, '') + ',' + end.replace(/-/g, '');
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    stats[row.ConversionTypeName] = row.Conversions;
  }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var header = data.shift();
  var col = header.indexOf(month) + 1;
  
  for(var z in data) {
    if(stats[data[z][1]]) {
      tab.getRange(parseInt(z,10)+2,col).setValue(stats[data[z][1]]);
      delete stats[data[z][1]];
    }
  }
  
  for(var name in stats) {
    var row = tab.getLastRow()+1;
    tab.getRange(row, 2).setValue(name);
    tab.getRange(row, col).setValue(stats[name]);
  }
}

function compileReport(id, url, contains, doesNotContain, labelName, filter, name, overall_flag, customer_id) {  
  var campaignMap = {};
  
  if(doesNotContain.length || labelName) {
    var iter = AdWordsApp.campaigns()
    
    if(labelName) {
      iter.withCondition('LabelNames CONTAINS_ANY ["' + labelName + '"]');
    }
    
    for(var z in doesNotContain) {
      iter.withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + doesNotContain[z] + '"');
    }
    
    iter = iter.get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
     
    var iter = AdWordsApp.shoppingCampaigns()
    
    if(labelName) {
      iter.withCondition('LabelNames CONTAINS_ANY ["' + labelName + '"]');
    }
    
    for(var z in doesNotContain) {
      iter.withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + doesNotContain[z] + '"');
    }
    
    iter = iter.get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
    
    var iter = AdWordsApp.videoCampaigns()
    
    if(labelName) {
      iter.withCondition('LabelNames CONTAINS_ANY ["' + labelName + '"]');
    }
    
    for(var z in doesNotContain) {
      iter.withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + doesNotContain[z] + '"');
    }
    
    iter = iter.get();
    while(iter.hasNext()) {
      campaignMap[iter.next().getName()] = 1; 
    }
  }
  
  
  //Logger.log(campaignMap);
  var dt = new Date();
  dt.setHours(12);
  
  var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'));
  if(now.getDate() < 3) {
    dt.setDate(0);
    var month = Utilities.formatDate(dt, 'GMT', 'MMM');
    var end = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    var start = end.substring(0,8) + '01';
    compileReportForMonth(id, start, end, month, url, dt, contains, campaignMap, filter, name, overall_flag, customer_id); 
  }
  
  var dt = new Date();
  dt.setHours(12);
  var month = Utilities.formatDate(new Date(), 'GMT', 'MMM');
  var end = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  var start = end.substring(0,8) + '01';
  compileReportForMonth(id, start, end, month, url, dt, contains, campaignMap, filter, name, overall_flag, customer_id);
}

function compileReportForMonth(id, start, end, month, url, dt, contains, campaignMap, filter, name, overall_flag, customer_id) {
  var ty = Utilities.formatDate(dt, 'GMT', 'yyyy');
  var ss = SpreadsheetApp.openByUrl(url);
  
  var tabName = ty;
  if(name) {
    tabName = name + ' - ' + tabName;
  }
  
  var tab = ss.getSheetByName(tabName);
  if(!tab) {
    //ss.setActiveSheet(ss.getSheetByName('Template'))
    var TEMP_NAME = TEMPLATE_TAB_NAME;
    if(overall_flag == 'Y') {
      TEMP_NAME += ' 2';
    }
    tab = SpreadsheetApp.openByUrl(TEMPLATE_URL).getSheetByName(TEMP_NAME).copyTo(ss);
    tab.setName(tabName);
    tab.showSheet();
  }
  
  var stats = {
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0,
    'AssistedConversions': 0, 'AssistedConversionValue': 0, 'Sessions': 0,
    'Clicks': 0, 'NewUsers': 0, 'OverallRevenue': 0, 'Bounces': 0
  };
  
  var filters = ['ga:medium==cpc;ga:source==google']
  if(contains) {
    filters.push('ga:campaign=@' + contains);
  }
  
  if(customer_id) {
    filters.push('ga:adwordsCustomerID==' + customer_id);
  }
  
  if(filter) {
    filters.push(filter);
  }
  
  //Logger.log(campaignMap);
  if(!Object.keys(campaignMap).length) {
    var optArgs = { 'filters': filters.join(';') };
    getDataFromAnalytics(id,stats,start,end,optArgs,overall_flag);
    getDataFromMCF(id,stats,start,end,contains,customer_id);
  } else {
    var optArgs = { 'dimensions': 'ga:campaign', 'filters': filters.join(';') };
    getCampaignDataFromAnalytics(id,stats,start,end,optArgs,campaignMap);
    getCampaignDataFromMCF(id,stats,start,end,campaignMap,contains,customer_id);
  }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var header = data.shift();
  var col = header.indexOf(month) + 1;
  
  //Logger.log(col);
  var rows = [[stats.Cost], [stats.ConversionValue], [stats.Conversions],
              [stats.AssistedConversions], [stats.AssistedConversionValue]];
  //Logger.log(rows);  
  tab.getRange(5, col, rows.length, 1).setValues(rows);
  
  if(overall_flag == 'Y') {
    var cr = stats.Clicks == 0 ? 0 : stats.Conversions / stats.Clicks;
    var pct = stats.Sessions == 0 ? 0 : stats.NewUsers / stats.Sessions;
    
    tab.getRange(11, col, 2, 1).setValues([[cr],[pct]]);
    tab.getRange(27, col).setValue(stats.OverallRevenue);
  }
}

function getDataFromAnalytics(PROFILE_ID,stats,FROM,TO,optArgs,overall_flag) {
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue,ga:sessions,ga:adClicks,ga:newUsers,ga:bounces",
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
    stats.Cost += parseFloat(rows[k][0]);
    stats.Conversions += parseInt(rows[k][1],10);
    stats.ConversionValue += parseFloat(rows[k][2]);
    stats.Sessions += parseInt(rows[k][3],10);  
    stats.Clicks += parseInt(rows[k][4], 10);
    stats.NewUsers += parseInt(rows[k][5], 10);
    stats.Bounces += parseInt(rows[k][6], 10);
  }
  
  if(overall_flag == 'Y') {
    var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                  // End-date (format yyyy-MM-dd).
          "ga:transactionRevenue");
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + PROFILE_ID);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var rows = resp.getRows();
    for(var k in rows) {
      stats.OverallRevenue += parseFloat(rows[k][0]);
    }
  }
}

function getDataFromMCF(PROFILE_ID,stats,FROM,TO,contains,customer_id) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  if(contains) {
    filters.push('mcf:adwordsCampaignPath=@' + contains); 
  }
  
  if(customer_id) {
    filters.push('mcf:adwordsCustomerID==' + customer_id); 
  }
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType',
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
    
    stats.AssistedConversions += parseInt(rows[k][2].primitiveValue,10);
    stats.AssistedConversionValue += parseFloat(rows[k][3].primitiveValue);    
  }
}

function getCampaignDataFromAnalytics(PROFILE_ID,stats,FROM,TO,optArgs,campaignMap) {
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue,ga:sessions,ga:adClicks,ga:bounces",
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
    var camp = rows[k][0];
    if(!campaignMap[camp]) { continue; }
    stats.Cost += parseFloat(rows[k][1]);
    stats.Conversions += parseInt(rows[k][2],10);
    stats.ConversionValue += parseFloat(rows[k][3]);
    stats.Sessions += parseInt(rows[k][4],10);    
    stats.Clicks += parseInt(rows[k][5],10);    
    stats.Bounces += parseInt(rows[k][6],10);    
  }
}

function getCampaignDataFromMCF(PROFILE_ID,stats,FROM,TO,campaignMap,contains,customer_id) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  if(contains) {
    filters.push('mcf:adwordsCampaignPath=@' + contains); 
  }
  
  if(customer_id) {
    filters.push('mcf:adwordsCustomerID==' + customer_id); 
  }
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:adwordsCampaignPath',
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
    var camp = JSON.parse(rows[k][1]).conversionPathValue[0].nodeValue;
    if(!campaignMap[camp]) { continue; }
    
    stats.AssistedConversions += parseInt(rows[k][2].primitiveValue,10);
    stats.AssistedConversionValue += parseFloat(rows[k][3].primitiveValue);    
  }
}



function compileWeeklyMonthlyReports(id, REPORT_URL,overall_flag) {
  var month = getAdWordsFormattedDate(1, 'MMMM yyyy');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';

  compileWMReportForDates(id, FROM, TO, month, 'Monthly', REPORT_URL,overall_flag);
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day - 1;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(sd, 'PST', 'yyyy-MM-dd');
  
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endDate =  Utilities.formatDate(ed, 'PST', 'yyyy-MM-dd');
  
  var key = start + ' - ' + end;
  compileWMReportForDates(id, startDate, endDate, key, 'Weekly', REPORT_URL,overall_flag);
  
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  var day = now.getDay();
  if(day == 0) { day = 7; }
  
  var diff = day + 6;
  var sd = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  sd.setDate(sd.getDate() - diff);
  var start = Utilities.formatDate(sd, 'PST', 'MMM d, yyyy');
  var startDate =  Utilities.formatDate(sd, 'PST', 'yyyy-MM-dd');
  
  var diff = day;
  var ed = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  ed.setDate(ed.getDate() - diff);
  var end = Utilities.formatDate(ed, 'PST', 'MMM d, yyyy');
  var endDate =  Utilities.formatDate(ed, 'PST', 'yyyy-MM-dd');
  
  var key = start + ' - ' + end;
  compileWMReportForDates(id, startDate, endDate, key, 'Weekly', REPORT_URL,overall_flag);
}

function compileWMReportForDates(id, FROM, TO, dateKey, REPORT_KEY, REPORT_URL,overall_flag) {
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(REPORT_KEY + ' Report');
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var rNum = -1;
  for(var k in data) {
    if(data[k][0] == dateKey) {
     rNum = parseInt(k,10) + 3; 
    }
    //outputData[data[k][0]] = [data[k][18],data[k][19],data[k][20],data[k][21],data[k][22],data[k][23]];
  }
  
  if(rNum == -1) { return; }
  
  var optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
  
  var googleStats = {
    'Sessions': 0, 'ConversionValue': 0, 'Cost': 0, 'Conversions': 0
  };
  
  getDataFromAnalytics(id,googleStats,FROM,TO,optArgs,overall_flag);
  
  var overallStats = {
    'Sessions': 0, 'ConversionValue': 0, 'Cost': 0, 'Conversions': 0
  };
  
  var optArgs = {};
  getDataFromAnalytics(id,overallStats,FROM,TO,optArgs,overall_flag);
  
  var row = {
    'PPCSessions':  googleStats['Sessions'], 'PPCRevenue':  googleStats['ConversionValue'],
    'TotalSessions':  overallStats['Sessions'], 'TotalRevenue':  overallStats['ConversionValue']
  }
  
  var out = [
    [row.TotalSessions, row.PPCSessions, row.TotalSessions == 0 ? 0 : round(row.PPCSessions / row.TotalSessions, 4),
    row.TotalRevenue, row.PPCRevenue, row.TotalRevenue == 0 ? 0 : round(row.PPCRevenue / row.TotalRevenue, 4)]
  ];
  
  
  sheet.getRange(rNum,19,out.length,out[0].length).setValues(out);
}


function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}