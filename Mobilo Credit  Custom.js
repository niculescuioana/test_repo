function main() {
    MccApp.accounts().withIds(['830-456-9344']).executeInParallel('run');
  }
  
  function run() {
    compileAdWordsReport();
    compileGAReport();
  }
  
  function compileAdWordsReport() {
    var URL = 'https://docs.google.com/spreadsheets/d/1Vo5mc1xbPSdXM_U8_ZeozEgbjtssbuh5l54KuEvgrXQ/edit#gid=1629213731'; 
    
    var dt = new Date();
    dt.setHours(12);
    var month = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM yyyy');
    var end = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    var start = end.substring(0,6) + '01';
    
    var DATE_RANGE = start + ',' + end;
    
    dt.setDate(1);
    dt.setDate(0);
    dt.setHours(12);
    
    var lm = Utilities.formatDate(dt, 'PST', 'MMM yyyy');
    
    var ss = SpreadsheetApp.openByUrl(URL);
    var tab = ss.getSheetByName(month);
    if(!tab) {
      ss.setActiveSheet(ss.getSheetByName(lm));
      tab = ss.duplicateActiveSheet();
      tab.setName(month);
    }
    
    tab.getRange(2,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
    
    var cols = [
      'Date',  'Cost', 'Impressions', 'Clicks', 'AverageCpc', 'Ctr', 'AveragePosition'
    ];
    
    var query = [
      'SELECT', cols.join(','), 'FROM', 'ACCOUNT_PERFORMANCE_REPORT',
      'DURING ', DATE_RANGE 
    ].join(' ');
    
    var map = {};
    var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
    while(rows.hasNext()) {
      var row = rows.next();
      map[row.Date] = {
        'Clicks': parseInt(row.Clicks,10),
        'Impressions': parseInt(row.Impressions,10),
        'Cost': parseFloat(row.Cost.toString().replace(/,/g, '')),
        'AverageCpc': parseFloat(row.AverageCpc),
        'AveragePosition': parseFloat(row.AveragePosition),
        'Ctr': row.Ctr,
        'Conversions': 0
      }
    }
    
    var cols = [
      'Date', 'ConversionTypeName', 'Conversions'
    ];
    
    var query = [
      'SELECT', cols.join(','), 'FROM', 'ACCOUNT_PERFORMANCE_REPORT',
      'where ConversionTypeName = "Aplicare Finalizata (All Web Site Data)"',
      'DURING', DATE_RANGE
    ].join(' ');
    
    var rows = AdWordsApp.report(query).rows();
    while(rows.hasNext()) {
      var row = rows.next();
      map[row.Date]['Conversions'] = parseFloat(row.Conversions);
    }
    
    var out = [];
    for(var date in map) {
      var row = map[date];
      out.push([date, row.Impressions, row.Clicks, row.Ctr, row.AverageCpc, row.Cost, row.Conversions,
                row.Conversions == 0 ? 0 : row.Cost / row.Conversions, row.AveragePosition]); 
    }
    
    if(out.length) {
      tab.getRange(2,1,out.length,out[0].length).setValues(out);
      tab.sort(1);
    }
  }
  
  function compileGAReport() {
    var URL = 'https://docs.google.com/spreadsheets/d/1Vo5mc1xbPSdXM_U8_ZeozEgbjtssbuh5l54KuEvgrXQ/edit#gid=1629213731'; 
    
    var dt = new Date();
    dt.setHours(12);
    var month = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM yyyy');
    var end = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
    var start = end.substring(0,8) + '01';
    
    dt.setDate(1);
    dt.setDate(0);
    dt.setHours(12);
    
    var lm = Utilities.formatDate(dt, 'PST', 'MMM yyyy');
    
    var ss = SpreadsheetApp.openByUrl(URL);
    var tab = ss.getSheetByName('Analytics: ' + month);
    if(!tab) {
      ss.setActiveSheet(ss.getSheetByName('Analytics: ' + month));
      tab = ss.duplicateActiveSheet();
      tab.setName('Analytics: ' + month);
    }
    
    tab.getRange(2,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
    
    var optArgs = { 'dimensions': 'ga:date,ga:sourceMedium', 'samplingLevel': 'HIGHER_PRECISION', 'max-results': 10000 };
    var results = Analytics.Data.Ga.get('ga:116736184',start,end,"ga:users,ga:newUsers,ga:sessions,ga:bounceRate,ga:pageviewsPerSession,ga:avgSessionDuration,ga:goal14ConversionRate,ga:goal14Completions",optArgs);
    var rows = results.getRows();
    for(var i in rows) {
         rows[i][0] = rows[i][0].substring(0,4) + '-' +  rows[i][0].substring(4,6) + '-' + rows[i][0].substring(6,8);
      rows[i][5] = rows[i][5]+'%';
      rows[i][8] = rows[i][8]+'%';
    }
    
    tab.getRange(2,1,rows.length,rows[0].length).setValues(rows).sort([{'column':1, 'ascending': true},{'column': 3, 'ascending': false}]);
    //tab.sort(1,false);
  }