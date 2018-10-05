function main() {
    MccApp.select(MccApp.accounts().withIds(["570-380-5545"]).get().next());
    
    var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1WGqrhpqQvCpySSbjLGDwqlDiMPgY7h4baqNJeMHVw2k/edit#gid=735465664';
    var ID = '104503053';
    
    var map = {
      'TV&Multimedia': 23,
      'Solutii_Mobile': 29,
      'Portabile': 35,
      'Electrocasnice': 41,
      'Others': 47,
      'Display': 53,
      'Brand_Only': 59,
      'MultiCategory': 65,
      'Mama Si Copilul': 71,
      'Ingrijire Personala': 77,
      'Curatenia Casei': 83,
      'Bricolaj': 89,
      'Componente PC': 95,
      'Imprimante': 101,
      'auto': 107,
      'Video Category': 113,
      'Monitoare': 119,
      'Electrocasnice - Frigorifice': 125,
      'Electrocasnice - Racire': 131,    
      'Sport & Fitness - Electrice': 137
    };
    
    for(var label in map) {
      compileReport(label, ID, REPORT_URL, map[label]);
    }
  }
  
  function compileReport(labelName, id, url, rowNum) {
    if(!AdWordsApp.labels().withCondition('Name = "' + labelName + '"').get().hasNext()) {
      return; 
    }
    
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
    
    if(!Object.keys(campaignMap).length) { return; }
    
    var dt = new Date();
    dt.setHours(12);
    
    var now = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy'));
    var month = Utilities.formatDate(dt, 'GMT', 'MMM yyyy');
    var date = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    compileReportForDay(id, date, month, url, dt, campaignMap, rowNum); 
    
    dt.setDate(dt.getDate()-1);
    var month = Utilities.formatDate(dt, 'GMT', 'MMM yyyy');
    var date = Utilities.formatDate(dt, 'GMT', 'yyyy-MM-dd');
    compileReportForDay(id, date, month, url, dt, campaignMap, rowNum); 
  }
  
  function compileReportForDay(id, date, month, url, dt, campaignMap, rowNum) {
    var ss = SpreadsheetApp.openByUrl(url);
    var tabName = 'AdWords: ' +month;
    var tab = ss.getSheetByName(tabName);
    if(!tab) {
      return;
    }
    
    var stats = {
      'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Sessions': 0
    };
    
    var optArgs = { 'dimensions': 'ga:campaign', 'filters': 'ga:medium==cpc;ga:source==google' };
    getCampaignDataFromAnalytics(id,stats,date,optArgs,campaignMap);
    
    
    var header = tab.getRange(2,1,1,tab.getLastColumn()).getValues();
    var col = dt.getDate() + 3;
    var rows = [
      [stats.Cost], [stats.ConversionValue], [stats.Conversions]
    ];
    
    tab.getRange(rowNum, col, rows.length, 1).setValues(rows);
  }
  
  function getCampaignDataFromAnalytics(PROFILE_ID,stats,date,optArgs,campaignMap) {
    var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          date,                 // Start-date (format yyyy-MM-dd).
          date,                  // End-date (format yyyy-MM-dd).
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
      var camp = rows[k][0];
      if(!campaignMap[camp]) { continue; }
      stats.Cost += parseFloat(rows[k][1]);
      stats.Conversions += parseInt(rows[k][2],10);
      stats.ConversionValue += parseFloat(rows[k][3]);
      stats.Sessions += parseInt(rows[k][4],10);    
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