function main() {
    MccApp.accounts().withIds(['524-985-9325']).executeInParallel('run');
  }
  
  function run() {
    var today = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    var URL = 'https://docs.google.com/spreadsheets/d/1Om73WyucRJqjsfsml2X42K2VHh5aNTWYVv4pmUaemDk/edit#gid=0';
    var query = [
      'SELECT ConversionTypeName, Date, HourOfDay, Conversions FROM ACCOUNT_PERFORMANCE_REPORT',
      'DURING 20180401,', today
    ].join(' ');
    
    var out = [];
    var rows = AdWordsApp.report(query).rows();
    while(rows.hasNext()){
      var row = rows.next();
      out.push([row.ConversionTypeName, row.Date, row.HourOfDay, row.Conversions]);
    }
    
    var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('AdWords');
    tab.getRange(2,1,out.length,out[0].length).setValues(out);
    tab.sort(2, false);
    
    compileAnalyticsReport(URL, 74497460, '2018-04-01', Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd'));
  }
  
  function compileAnalyticsReport(URL, PROFILE_ID, FROM, TO) {
   var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                  // End-date (format yyyy-MM-dd).
          "ga:goalCompletionsAll",
          {
          'dimensions': 'ga:goalCompletionLocation,ga:goalPreviousStep1,ga:goalPreviousStep2,ga:goalPreviousStep3',
          'segment': 'users::condition::ga:source=@google;ga:medium=@cpc'
          });
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + PROFILE_ID);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var out = resp.getRows();
     
    var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Analytics');
    tab.getRange(2,1,out.length,out[0].length).setValues(out);
    tab.sort(5, false);  
  }