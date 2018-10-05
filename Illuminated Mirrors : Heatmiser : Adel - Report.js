/*************************************************
* Illuminated Mirrors - Monthly Reporting
* @version: 1.0
* @author: Naman Jindal (nj.itprof@gmail.com)
***************************************************/

var MONTHLY_TAB_NAME = 'Google & Bing Update';

var START_DATE = '20151116';
var LAST_N_MONTHS = 2;

var PROFILE_ID = 26474681;

function main() {
  var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1qJjmvuHBTVdGF5kxGWwDa2_3xuRMosQ2h5n8o7W4YA8/edit';
  runScript("Illuminated Mirrors", REPORT_URL);
  
  REPORT_URL = 'https://docs.google.com/spreadsheets/d/1Rr5bkLGJaH7Vmb8ubOin3DKJ5mNuL_imG7NdRBYnnhw/edit';
  compileHeatmiserAdelReport('Heatmiser', REPORT_URL, 1373929);
  
  REPORT_URL = 'https://docs.google.com/spreadsheets/d/1vSLqcNVdgpNUwvSP2XW9y__FY3F3OKH6eKD1QJoLMBs/edit';
  compileHeatmiserAdelReport('Adel Direct', REPORT_URL, 115467935);
}

function compileHeatmiserAdelReport(ACC_NAME, REPORT_URL, ID) {
  MccApp.select(MccApp.accounts().withCondition('Name = "'+ ACC_NAME +'"').get().next());
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  date.setDate(date.getDate()-1);
  date.setHours(12);
  
  var yesterday = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  var yestLy = parseInt(yesterday.substring(0,4),10)-1 + '' + yesterday.substring(4,10);
  var yearStart = yesterday.substring(0,4) + '-01-01';
  
  var thisMonthStart = yesterday.substring(0,8) + '01';
  var yestMtdLyStart = yestLy.substring(0,8) + '01';
  var lastYearStart = yestLy.substring(0,4) + '-01-01';
  
  
  date.setMonth(date.getMonth()-1);
  var lastMonthEnd = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  var lastMonthStart = lastMonthEnd.substring(0,8) + '01';
  
  date.setMonth(date.getMonth()-10);
  date.setDate(0);
  
  var lastYearMonthEnd = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');;
  var lastYearMonthStart = lastYearMonthEnd.substring(0,8) + '01';
  
  
  var statsTM = getDataForRevenueReport(ID,thisMonthStart,yesterday,false);
  var statsLY = getDataForRevenueReport(ID,lastYearMonthStart,lastYearMonthEnd,false);  
  
  var statsYTD = getDataForRevenueReport(ID,yearStart,yesterday,false);  
  var statsLYTD = getDataForRevenueReport(ID,lastYearStart,yestLy,false);  
  var statsMTDLY = getDataForRevenueReport(ID,yestMtdLyStart,yestLy,false);  
  
  var statsLMTD = getDataForRevenueReport(ID,lastMonthStart,lastMonthEnd,false);  
  
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  now.setDate(now.getDate()-1);
  now.setDate(1);
  now.setDate(0);
  
  var END = Utilities.formatDate(now, 'PST', 'yyyy-MM-dd');
  var START = END.substring(0,8) + '01'; 
  var statsLM = getDataForRevenueReport(ID,START,END,false);  
  
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Summary');
  
  tab.getRange(21, 2, 2, 2).setValues([[statsTM.Conversions, statsMTDLY.Conversions],
                                       [statsTM.AssistedConversions, statsMTDLY.AssistedConversions]]);
  
  tab.getRange(21, 6, 2, 4).setValues([[statsLMTD.Conversions, '', statsLM.Conversions, statsLY.Conversions],
                                       [statsLMTD.AssistedConversions, '', statsLM.AssistedConversions, statsLY.AssistedConversions]]);
  
  tab.getRange(21, 12, 2, 2).setValues([[statsYTD.Conversions, statsLYTD.Conversions],
                                        [statsYTD.AssistedConversions, statsLYTD.AssistedConversions]]);
  
  tab.getRange(25, 2, 2, 2).setValues([[statsTM.ConversionValue, statsMTDLY.ConversionValue],
                                       [statsTM.AssistedConversionValue, statsMTDLY.AssistedConversionValue]]);
  
  tab.getRange(25, 6, 2, 4).setValues([[statsLMTD.ConversionValue, '', statsLM.ConversionValue, statsLY.ConversionValue],
                                       [statsLMTD.AssistedConversionValue, '', statsLM.AssistedConversionValue, statsLY.AssistedConversionValue]]);
  
  tab.getRange(25, 12, 2, 2).setValues([[statsYTD.ConversionValue, statsLYTD.ConversionValue],
                                        [statsYTD.AssistedConversionValue, statsLYTD.AssistedConversionValue]]);                                       
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  
  if(date.getDate() == 1) {
    compileHeatmiserAdelHistoricalReport(ID, REPORT_URL, 1);
  }
  
  compileHeatmiserAdelHistoricalReport(ID, REPORT_URL, 0);
}


function runScript(ACC_NAME, REPORT_URL) {
  try {
    var mccAccount = AdWordsApp.currentAccount();
    MccApp.select(MccApp.accounts().withCondition('Name = "'+ ACC_NAME +'"').get().next());
    compileMonthlyReport(REPORT_URL);
    compilRevenueReport(REPORT_URL);
  } catch (e) {
    Logger.log(e);
  }
}

function readBingStats() {
  var map = {};
  
  var folder = DriveApp.getFolderById('0BwnikHB3eS37SnU2c3h3Y0wyNGs');
  var file = folder.getFilesByName('Illuminati_Monthly_TM.csv').next();
  var data = Utilities.parseCsv(file.getBlob().getDataAsString());
  data.shift();
  
  for(var x in data) {
    map[data[x][0]] = {
      'ConversionValue': 0/*parseFloat(data[x][8])*/, 'Cost': parseFloat(data[x][5]), 
      'Conversions': 0/*parseInt(data[x][7],10)*/, 'CostPerConversion': 0, 'ROAS': 0
    }
  }
  
  var file = folder.getFilesByName('Illuminati_Monthly_6M.csv').next();
  var data = Utilities.parseCsv(file.getBlob().getDataAsString());
  data.shift();
  
  for(var x in data) {
    map[data[x][0]] =  {
      'ConversionValue': 0/*parseFloat(data[x][8])*/, 'Cost': parseFloat(data[x][5]), 
      'Conversions': 0/*parseInt(data[x][7],10)*/, 'CostPerConversion': 0, 'ROAS': 0
    }
  }
  
  return map;
}

function compileMonthlyReport(REPORT_URL) {
  
  var analyticsStatsMap = readStatsFromAnalytics();
  var bingStatsMap = readBingStats();
  
  var sheet = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(MONTHLY_TAB_NAME);
  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var outputData = [];
  for(var k in data) {
    outputData.push(data[k]);
  }
  
  for(var key in analyticsStatsMap['google']) {
    var found = false;
    
    var row = analyticsStatsMap['google'][key];
    if(!row) {
      row = { 'ConversionValue': 0, 'Cost': 0, 'Conversions': 0, 'CostPerConversion': 0, 'ROAS': 0 }
    }
    
    var bingRow = bingStatsMap[key];
    if(!bingRow) {
      bingRow = { 'ConversionValue': 0, 'Cost': 0, 'Conversions': 0, 'CostPerConversion': 0, 'ROAS': 0 }
    }
    
    var bingAnalyticsRow = analyticsStatsMap['bing'][key];
    if(!bingAnalyticsRow) {
      bingAnalyticsRow = { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0 }
    }
    
    bingRow.Cost += bingAnalyticsRow.Cost;
    bingRow.Conversions += bingAnalyticsRow.Conversions;
    bingRow.ConversionValue += bingAnalyticsRow.ConversionValue;
    
    var totalsRow = {
      'ConversionValue': 0, 'Cost': 0, 'Conversions': 0, 'CostPerConversion': 0, 'ROAS': 0
    }
    
    totalsRow.Cost = bingRow.Cost + row.Cost;
    totalsRow.ConversionValue = bingRow.ConversionValue + row.ConversionValue;
    totalsRow.Conversions = bingRow.Conversions + row.Conversions;
    
    totalsRow.CostPerConversion = totalsRow.Conversions == 0 ? 0 : round(totalsRow.Cost/totalsRow.Conversions,2);
    totalsRow.ROAS = totalsRow.Cost == 0 ? 0 : round(totalsRow.ConversionValue/totalsRow.Cost,2);
    
    bingRow.CostPerConversion = bingRow.Conversions == 0 ? 0 : round(bingRow.Cost/bingRow.Conversions,2);
    bingRow.ROAS = bingRow.Cost == 0 ? 0 : round(bingRow.ConversionValue/bingRow.Cost,2);
    
    row.CostPerConversion = row.Conversions == 0 ? 0 : round(row.Cost/row.Conversions,2);
    row.ROAS = row.Cost == 0 ? 0 : round(row.ConversionValue/row.Cost,2);
    
    for(var k in outputData) {
      if(outputData[k][0] == key) {
        outputData[k] = [key, row.ConversionValue, row.Conversions, row.Cost, row.CostPerConversion, row.ROAS,
                         bingRow.ConversionValue, bingRow.Conversions, bingRow.Cost, bingRow.CostPerConversion, bingRow.ROAS,
                         totalsRow.ConversionValue, totalsRow.Conversions, totalsRow.Cost, totalsRow.CostPerConversion, totalsRow.ROAS];
        found = true;
        break;
      }
    }
    
    if(!found) {
      outputData.push([key, row.ConversionValue, row.Conversions, row.Cost, row.CostPerConversion, row.ROAS,
                       bingRow.ConversionValue, bingRow.Conversions, bingRow.Cost, bingRow.CostPerConversion, bingRow.ROAS,
                       totalsRow.ConversionValue, totalsRow.Conversions, totalsRow.Cost, totalsRow.CostPerConversion, totalsRow.ROAS]);
    } 
  }
  
  if(outputData.length > 0) {
    sheet.getRange(3,1,outputData.length,outputData[0].length).setValues(outputData);
  }
  
}

function readStatsFromAnalytics() {
  var keys = {};
  var date  = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
  keys[Utilities.formatDate(date, 'PST', 'yyyyMM')] = Utilities.formatDate(date, 'PST', 'MMMM yyyy') ;
  
  date.setDate(0);
  date.setHours(12);
  keys[Utilities.formatDate(date, 'PST', 'yyyyMM')] = Utilities.formatDate(date, 'PST', 'MMMM yyyy') ;
  
  while(LAST_N_MONTHS > 1) {
    date.setDate(0);
    LAST_N_MONTHS--;
    keys[Utilities.formatDate(date, 'PST', 'yyyyMM')] = Utilities.formatDate(date, 'PST', 'MMMM yyyy') ;
  }
  
  date.setDate(1);
  date.setHours(12);
  
  var FROM = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
  var TO = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  
  var optArgs = { 'dimensions': 'ga:yearMonth', 'filters': 'ga:medium==cpc;ga:source==google' };
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  
  var results = {
    'google': {},
    'bing': {}
  };
  
  for(var k in rows) {
    var month = keys[rows[k][0]];
    results['google'][month] = { 
      'Cost': parseInt(rows[k][1],10),
      'Conversions': parseInt(rows[k][2],10),
      'ConversionValue': parseFloat(rows[k][3])
    }
  }
  
  
  var optArgs = { 'dimensions': 'ga:yearMonth', 'filters': 'ga:medium==cpc;ga:source==bing' };
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + id);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    var month = keys[rows[k][0]];
    results['bing'][month] = { 
      'Cost': parseInt(rows[k][1],10),
      'Conversions': parseInt(rows[k][2],10),
      'ConversionValue': parseFloat(rows[k][3])
    }
  }
  
  return results;
}

function compilRevenueReport(REPORT_URL) {
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8)+'01';
  
  var statsNow = getDataForRevenueReport(PROFILE_ID, FROM, TO, true);
  
  var LY = parseInt(FROM.substring(0,4),10)-1;
  FROM = LY+''+FROM.substring(4,10);
  TO = LY+''+TO.substring(4,10);
  
  var statsPrev = getDataForRevenueReport(PROFILE_ID, FROM, TO, true);
  
  var rowNum = parseInt(getAdWordsFormattedDate(0,'MM'),10) + 2;
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Target Dashboard');
  tab.getRange(rowNum,2,1,2).setValues([[statsPrev.ConversionValue, statsPrev.AssistedConversionValue]]);
  tab.getRange(rowNum,5,1,2).setValues([[statsNow.ConversionValue, statsNow.AssistedConversionValue]]);   
}

function getDataForRevenueReport(ID, FROM, TO, includeBing) {
  
  var stats = {
    'Conversions': 0, 'ConversionValue': 0, 
    'AssistedConversions': 0, 'AssistedConversionValue': 0
  }
  
  var optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      resp = Analytics.Data.Ga.get(
        'ga:'+ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:transactionRevenue,ga:transactions",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  for(var k in rows) {
    stats.ConversionValue += parseFloat(rows[k][0]);
    stats.Conversions += parseInt(rows[k][1], 10);
  }
  
  if(includeBing) {
    var optArgs = { 'filters': 'ga:medium==cpc;ga:source==bing' };
    var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        resp = Analytics.Data.Ga.get(
          'ga:'+ID,              // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                   // End-date (format yyyy-MM-dd).
          "ga:transactionRevenue",
          optArgs);
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + ID);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var rows = resp.getRows();
    for(var k in rows) {
      stats.ConversionValue += parseFloat(rows[k][0]);
      stats.Conversions += parseInt(rows[k][1], 10);    
    }
  }
  
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType',
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+ID,      // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                  // End-date (format yyyy-MM-dd).
    "mcf:totalConversionValue,mcf:totalConversions",
    optArgs
  );
  
  var rows = results.rows;
  
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    stats.AssistedConversionValue += parseFloat(rows[k][2].primitiveValue);    
    stats.AssistedConversions += parseInt(rows[k][3].primitiveValue, 10);    
  }
  
  return stats;
}



function compileHeatmiserAdelHistoricalReport(ID, REPORT_URL, N) {
  var TO = getAdWordsFormattedDate(N, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8)+'01';
  
  var googleStatsNow = AdWordsApp.currentAccount().getStatsFor(FROM.replace(/-/g, ''), TO.replace(/-/g, ''))
  var statsNow = getDataForRevenueReport(ID, FROM, TO, false);
  
  var LY = parseInt(FROM.substring(0,4),10)-1;
  FROM = LY+''+FROM.substring(4,10);
  TO = LY+''+TO.substring(4,10);
  
  var googleStatsPrev = AdWordsApp.currentAccount().getStatsFor(FROM.replace(/-/g, ''), TO.replace(/-/g, ''))
  var statsPrev = getDataForRevenueReport(ID, FROM, TO, false);
  
  var month = getAdWordsFormattedDate(N, 'MMMM');
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('2018 vs 2017');
  var header = tab.getDataRange().getValues()[0];
  var col = header.indexOf(month) + 1;
  
  tab.getRange(5, col, 5, 2).setValues([
    [googleStatsNow.getCost(), googleStatsPrev.getCost()],
    [googleStatsNow.getClicks(), googleStatsPrev.getClicks()],
    [googleStatsNow.getImpressions(), googleStatsPrev.getImpressions()],
    [googleStatsNow.getAverageCpc(), googleStatsPrev.getAverageCpc()],
    [googleStatsNow.getCtr(), googleStatsPrev.getCtr()]
  ]);
  
  tab.getRange(14, col, 2, 2).setValues([[statsNow.Conversions, statsPrev.Conversions],
                                        [statsNow.AssistedConversions, statsPrev.AssistedConversions]]);
  tab.getRange(19, col, 2, 2).setValues([[statsNow.ConversionValue, statsPrev.ConversionValue],
                                  [statsNow.AssistedConversionValue, statsPrev.AssistedConversionValue]]);
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function info(msg) {
  var time = Utilities.formatDate(new Date(),AdWordsApp.currentAccount().getTimeZone(),'MMM dd, yyyy HH:mm:ss.SSS');
  Logger.log(time + ' - ' + msg);
}  


function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}