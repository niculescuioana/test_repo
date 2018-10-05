var CONFIG = {
    '675360': 'https://docs.google.com/spreadsheets/d/1qkycIs_GDWz2RZHuGkpNU0gV0PXpqOHcJr9ic0dI7kc/edit'
  }
  
  function main() {
    for(var id in CONFIG) {
      compileReport(id, CONFIG[id]); 
    }
  }
  
  function compileReport(id, url) {
    var dt = new Date();
    dt.setHours(12);
    
    var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'));
    if(now.getDate() < 3) {
    dt.setDate(0);
      var month = Utilities.formatDate(dt, 'GMT', 'MMM');
      var monthFull = Utilities.formatDate(dt, 'GMT', 'MMMM');
      var year = Utilities.formatDate(dt, 'GMT', 'yyyy');
      var end = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
      var start = end.substring(0,8) + '01';
      compileReportForMonth(id, start, end, month, monthFull, year, url, dt); 
    }
    
    var dt = new Date();
    dt.setHours(12);
    var month = Utilities.formatDate(new Date(), 'GMT', 'MMM');
    var monthFull = Utilities.formatDate(new Date(), 'GMT', 'MMMM');
    var year = Utilities.formatDate(new Date(), 'GMT', 'yyyy');
    var end = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
    var start = end.substring(0,8) + '01';
    compileReportForMonth(id, start, end, month, monthFull, year, url, dt);
  }
  
  function compileReportForMonth(id, start, end, month, monthFull, year, url, dt) {
    var ty = Utilities.formatDate(dt, 'GMT', 'yyyy');
    var ss = SpreadsheetApp.openByUrl(url);
    var tab = ss.getSheetByName(ty);
    if(!tab) {
      ss.setActiveSheet(ss.getSheetByName('Template'))
      tab = ss.duplicateActiveSheet(); 
      tab.setName(ty);
      tab.showSheet();
    }
    
    var stats = {
      'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0,
      'OverallConversionValue': 0, 'AssistedConversions': 0, 'AssistedConversionValue': 0
    };
    
    var optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
    getDataFromAnalytics(id,stats,start,end,optArgs);
    //getDataFromMCF(id,stats,start,end);
    
    var data = tab.getDataRange().getValues();
    data.shift();
    
    var header = data.shift();
    var col = header.indexOf(month) + 1;
    
    stats.CPC = stats.Clicks == 0 ? 0 : round(stats.Cost / stats.Clicks, 2);
    
    tab.getRange(11, col).setValue(stats.Cost);
    tab.getRange(17, col).setValue(stats.Clicks);
    tab.getRange(24, col).setValue(stats.ConversionValue);
    tab.getRange(31, col).setValue(stats.Conversions);
    tab.getRange(55, col).setValue(stats.OverallConversionValue);
    
    var bingStats = {
      'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0,
      'OverallConversionValue': 0, 'AssistedConversions': 0, 'AssistedConversionValue': 0
    };
    
    var optArgs = { 'filters': 'ga:medium==cpc;ga:source==bing' };
    getDataFromAnalytics(id,bingStats,start,end,optArgs);
    
    addBingCost(monthFull, year, bingStats);
    
    tab.getRange(12, col).setValue(bingStats.Cost);
    tab.getRange(18, col).setValue(bingStats.Clicks);  
    tab.getRange(25, col).setValue(bingStats.ConversionValue);
    tab.getRange(32, col).setValue(bingStats.Conversions);
  }
  
  function addBingCost(month, year, bingStats) {
    var url = 'https://docs.google.com/spreadsheets/d/1Fpyuykm5xeXnJrS6C_5VhWFUX049vjANGEGyiUShp9I/edit';
    var sheet = SpreadsheetApp.openByUrl(url).getSheetByName(month);
    if(!sheet) {
      sheet = SpreadsheetApp.openByUrl(url).getSheetByName(month + ' ' + year); 
    }
    
    if(!sheet) { return; }
    
    var data = sheet.getDataRange().getValues();
    data.shift();
    for(var z in data) {
      if(data[z][0] == 'X1544060') {
        bingStats.Cost += data[z][5];
        bingStats.Clicks += data[z][4];
      }
    }
  }
  
  function getDataFromAnalytics(PROFILE_ID,stats,FROM,TO,optArgs) {
    var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                  // End-date (format yyyy-MM-dd).
          "ga:adCost,ga:adClicks,ga:transactions,ga:transactionRevenue",
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
      stats.Clicks += parseFloat(rows[k][1]);
      stats.Conversions += parseInt(rows[k][2],10);
      stats.ConversionValue += parseFloat(rows[k][3]);
    }
    
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                  // End-date (format yyyy-MM-dd).
          "ga:adCost,ga:adClicks,ga:transactions,ga:transactionRevenue");
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + PROFILE_ID);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var rows = resp.getRows();
    for(var k in rows) {
      stats.OverallConversionValue += parseFloat(rows[k][3]);
    }
  }
  /*
  function getDataFromMCF(PROFILE_ID,stats,FROM, TO) {
    var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                   'mcf:conversionType==Transaction'];
    
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
  */
  
  function getAdWordsFormattedDate(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
  }
  
  function round(num,n) {    
    return +(Math.round(num + "e+"+n)  + "e-"+n);
  }