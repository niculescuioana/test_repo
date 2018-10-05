function main() {
    MccApp.accounts().withCondition('Name = "Barriers Direct"').executeInParallel('run');
  }
  
  function run() {
    var DATE_RANGE_MAP = getDateRange(13);
    
    var header = [''];
    var initMap = {}, map = {};
    for(var month in DATE_RANGE_MAP.MONTHS) {
      initMap[month] = { 'clicks': 0, 'conversions': 0, 'cost': 0 }; 
      map[month] = { 'clicks': 0, 'conversions': 0, 'cost': 0 }; 
      header.push(month);
    }
    
    var stats = {
      'Barriers': JSON.parse(JSON.stringify(initMap)),
      'Bike Parking': JSON.parse(JSON.stringify(initMap)),
      'Bollards': JSON.parse(JSON.stringify(initMap)),
      'Car Security': JSON.parse(JSON.stringify(initMap)),
      'Crowd Control': JSON.parse(JSON.stringify(initMap)),
      'Display': JSON.parse(JSON.stringify(initMap)),
      'Other': JSON.parse(JSON.stringify(initMap)),
      'Parking Posts': JSON.parse(JSON.stringify(initMap)),
      'Warehouse': JSON.parse(JSON.stringify(initMap))
    };
    
    var OPTIONS = { 'includeZeroImpressions' : false };
    var cols = ['MonthOfYear','Year','Impressions','Clicks','Conversions','Cost','AverageCpc','CostPerConversion'];
    var report = 'ACCOUNT_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'during',DATE_RANGE_MAP.RANGE].join(' ');
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      row.Month = row.MonthOfYear + ' ' + row.Year;
      map[row.Month].clicks += parseInt(row.Clicks, 10);
      map[row.Month].cost += parseFloat(row.Cost.toString().replace(/,/g,''));
      map[row.Month].conversions += parseInt(row.Conversions.toString().replace(/,/g,''), 10);
    }
    
    var cols = ['MonthOfYear','Year','Labels','CampaignId',
                'Impressions','Clicks','Conversions','Cost','AverageCpc','CostPerConversion'];
    var report = 'CAMPAIGN_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'during',DATE_RANGE_MAP.RANGE].join(' ');
   
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      row.Month = row.MonthOfYear + ' ' + row.Year;
      if(row.Labels == '--') { continue; }
      
      var labels = JSON.parse(row.Labels);
      for(var z in labels) {
        if(!stats[labels[z]]) { continue; }
        stats[labels[z]][row.Month].clicks += parseInt(row.Clicks, 10);
        stats[labels[z]][row.Month].cost += parseFloat(row.Cost.toString().replace(/,/g,''));
        stats[labels[z]][row.Month].conversions += parseInt(row.Conversions.toString().replace(/,/g,''), 10);
      }
    }
    
    var output = [header];
    var url = 'https://docs.google.com/spreadsheets/d/1O3Qcsv6hDq8UiSJAYcU0PFMUxZUNgyuwJEFkxRip8Cs/edit';
    var ss = SpreadsheetApp.openByUrl(url);
    
    for(var label in stats) {
      var sheet = ss.getSheetByName(label);
      /*if(sheet) {
      ss.deleteSheet(sheet);
      sheet = ss.getSheetByName('Template').copyTo(ss);
      sheet.setName(label);
      }*/
      
      if(!sheet) {
        sheet =  ss.getSheetByName('Template').copyTo(ss);
        sheet.setName(label);
      }
      
      sheet.showSheet();
      
      var row = stats[label];
      
      var out = [];
      for(var month in row) {
        out.push([month, row[month].cost]);
      }
      
      sheet.getRange(3,2,out.length,2).setValues(out).sort({column: 2, ascending: true});
      
      var out = [];
      for(var month in row) {
        out.push([month, row[month].conversions]);
      }
      
      sheet.getRange(20,2,out.length,2).setValues(out).sort({column: 2, ascending: true});
      
      
      var out = [], outRow = [label];
      for(var month in row) {
        row[month].cpa = row[month].conversions == 0 ? 0 : row[month].cost / row[month].conversions
        out.push([month, row[month].cpa]);
        outRow.push(row[month].cost/map[month].cost);
      }
      
      sheet.getRange(37,2,out.length,2).setValues(out).sort({column: 2, ascending: true});
      
      output.push(outRow);
    }
    
    sheet.getRange(37,2,out.length,2).setValues(out).sort({column: 2, ascending: true});
      
    var tab = ss.getSheetByName('Summary');
    tab.getRange(2,1,output.length,output[0].length).setValues(output);
  }
  
  function getDateRange(LAST_N_MONTHS) {
    
    var map = { 'RANGE': '', 'MONTHS': {} };
    var END_DATE = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
    
    var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
    var TO = getAdWordsFormattedDate(0, 'yyyyMMdd');
    
    date.setDate(1);
    map['MONTHS'][Utilities.formatDate(date, 'PST', 'MMMM yyyy')] = 1;
    
    var N = 0;
    while(N < LAST_N_MONTHS) {
      date.setDate(0);
      map['MONTHS'][Utilities.formatDate(date, 'PST', 'MMMM yyyy')] = 1;
      
      date.setDate(1);
      N++;
    }
    
    var FROM = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
    
    map.RANGE = FROM + ',' + TO;
    
    return map;
  }
  
  function getAdWordsFormattedDate(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
  }