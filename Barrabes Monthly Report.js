

var PROFILE_ID_MAP = {
    'Barrabes Sweden': '154873873',
    'Barrabes Austria': '168926271', //'111424238',
    'Barrabes France': '161312664', //'111400577',
    'Barrabes Belgium': '161312664', //111400577',
    'Barrabes Germany': '168926271', //'111424238',
   /* 'Barrabes Australia': '169706021', //'111424238',
    'Barrabes Canada': '169706021',
    'Barrabes USA': '169706021',*/
    'Barrabes UK': '156971212',
    'Barrabes Spain': '111416518',
    'Others': '169706021'
  }
  
  
  function main() {
    var iter = MccApp.accounts()
    .withCondition('Name CONTAINS_IGNORE_CASE "barrabes"')
    //.withCondition('Name CONTAINS_IGNORE_CASE "usa"')
    .withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "old"')
    //.withLimit(1)
    .executeInParallel('runScript');
  }
  
  function runScript() {
    var REPORT_URL = 'https://docs.google.com/spreadsheets/d/16wmo0U9Wwr2J4TO0RVXqb7ujGoGmYNbTgFPJVNfftNQ/edit';
    compileMonthlyReport(REPORT_URL);
    SpreadsheetApp.flush();
    
    var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1pTX6BmoNZXlealfHcYhZBSHXCP3OHRPicpG_CQIdfDM/edit';
    compileYoYReport(REPORT_URL);
  }
  
  function compileYoYReport(REPORT_URL) {
    var name = AdWordsApp.currentAccount().getName().replace(' Ads', '');
    var ss = SpreadsheetApp.openByUrl(REPORT_URL);
    var tab = ss.getSheetByName(name);
    if(!tab) {
      ss.setActiveSheet(ss.getSheetByName('Template'))
      tab = ss.duplicateActiveSheet();
      tab.setName(name);
    }
    
    var header = tab.getDataRange().getValues()[0];
    
    var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
    date.setHours(12);
    date.setDate(date.getDate()-1);
    //date.setDate(1);
    //date.setDate(0);
    
    var endDate = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
    
    date.setDate(1);
    var startDate = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
    
    var statsMap = getStatsFromAdWords(startDate, endDate);
    addAnalyticsStats(name, statsMap, startDate, endDate);
    
    for(var month in statsMap) {
      var out = [];
      var index = header.indexOf(month.split(' ')[0]) + 2;
      for(var week in statsMap[month]) {
        
        statsMap[month][week].CPC = statsMap[month][week].Clicks == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Clicks, 2); 
        statsMap[month][week].CPA = statsMap[month][week].Conversions == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Conversions, 2); 
        statsMap[month][week].ROAS = statsMap[month][week].Cost == 0 ? 0 : round(statsMap[month][week].ConversionValue / statsMap[month][week].Cost, 2);      
        
        out.push([statsMap[month][week].Clicks]);
        out.push([statsMap[month][week].CPC]);
        out.push([statsMap[month][week].Cost]);
        out.push([statsMap[month][week].Conversions]);
        out.push([statsMap[month][week].ConversionValue]);
        out.push([statsMap[month][week].ROAS]);
        out.push([statsMap[month][week].CPA]);
        
        break;
      }
      
      tab.getRange(3, index, out.length, out[0].length).setValues(out);
    }
    
    
    var year = parseInt(endDate.split('-')[0], 10) - 1;
    endDate = year + endDate.substring(4, 10);
    startDate = year + startDate.substring(4, 10);
    
    var statsMap = getStatsFromAdWords(startDate, endDate);
    addAnalyticsStats(name, statsMap, startDate, endDate);
    
    for(var month in statsMap) {
      var out = [];
      var index = header.indexOf(month.split(' ')[0]) + 1;
      for(var week in statsMap[month]) {
        
        statsMap[month][week].CPC = statsMap[month][week].Clicks == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Clicks, 2); 
        statsMap[month][week].CPA = statsMap[month][week].Conversions == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Conversions, 2); 
        statsMap[month][week].ROAS = statsMap[month][week].Cost == 0 ? 0 : round(statsMap[month][week].ConversionValue / statsMap[month][week].Cost, 2);      
        
        out.push([statsMap[month][week].Clicks]);
        out.push([statsMap[month][week].CPC]);
        out.push([statsMap[month][week].Cost]);
        out.push([statsMap[month][week].Conversions]);
        out.push([statsMap[month][week].ConversionValue]);
        out.push([statsMap[month][week].ROAS]);
        out.push([statsMap[month][week].CPA]);
        
        break;
      }
      
      tab.getRange(3, index, out.length, out[0].length).setValues(out);
    }
    
    SpreadsheetApp.flush();
    
  }
  
  
  
  
  
  
  
  
  function compileMonthlyReport(REPORT_URL) {
    var name = AdWordsApp.currentAccount().getName().replace(' Ads', '');
    var ss = SpreadsheetApp.openByUrl(REPORT_URL);
    var tab = ss.getSheetByName(name);
    if(!tab) {
      info('Missing Tab');
      return;
      ss.setActiveSheet(ss.getSheetByName('Template'))
      tab = ss.duplicateActiveSheet();
      tab.setName(name);
    }
    
    var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
    date.setHours(12);
    
    date.setDate(0);
    date.setDate(1);
    
    var endDate = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
    var startDate = Utilities.formatDate(date, 'PST', 'yyyy-MM-dd');
    
    
    var data = tab.getDataRange().getValues();
    
    var indexMap = {};
    for(var x in data) {
      if(!data[x][0]) { continue; }
      indexMap[data[x][0]] = parseInt(x,10) + 1;
    }
    
    var statsMap = getStatsFromAdWords(startDate, endDate);
    //Logger.log(JSON.stringify(statsMap));
    addAnalyticsStats(name, statsMap, startDate, endDate);
    
    for(var month in statsMap) {
      var out = [];
      for(var week in statsMap[month]) {
        statsMap[month][week].CPC = statsMap[month][week].Clicks == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Clicks, 2); 
        statsMap[month][week].CPA = statsMap[month][week].Conversions == 0 ? 0 : round(statsMap[month][week].Cost / statsMap[month][week].Conversions, 2); 
        statsMap[month][week].ROAS = statsMap[month][week].Cost == 0 ? 0 : round(statsMap[month][week].ConversionValue / statsMap[month][week].Cost, 2);      
        var row = [week, statsMap[month][week].Clicks, statsMap[month][week].CPC, 
                   statsMap[month][week].Cost, statsMap[month][week].Conversions, 
                   statsMap[month][week].ConversionValue, statsMap[month][week].ROAS, statsMap[month][week].CPA];
        out.push(row);
      }
      
      var rowNum = indexMap[month];
      if(!rowNum) {
        tab.insertRowsBefore(2, 7);
        rowNum = 2;
      }
      
      tab.getRange(rowNum, 1, out.length, out[0].length).setValues(out);
    }
    
    tab.getRange(2, 1, 1, tab.getLastColumn()).setFontWeight('bold');
    tab.getRange(3, 1, 5, tab.getLastColumn()).setFontWeight('normal');  
  }
  
  function addAnalyticsStats(name, map, startDate, endDate) {
    //var map = {};
    var ID = PROFILE_ID_MAP[name];
    if(!ID) {
      ID =  PROFILE_ID_MAP['Others'];
    }
    
    var optArgs = { 
      'dimensions': 'ga:date,ga:month,ga:year,ga:adwordsCustomerID', 
      'filters': 'ga:medium==cpc;ga:source==google' 
    };
    
    var attempts = 3;
    var results;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        results = Analytics.Data.Ga.get(
          'ga:'+ID,              // Table id (format ga:xxxxxx).
          startDate,                 // Start-date (format yyyy-MM-dd).
          endDate,                   // End-date (format yyyy-MM-dd).
          "ga:transactionRevenue,ga:impressions,ga:adClicks,ga:adCost,ga:transactions",
          optArgs);
        
        break;
      } catch(ex) {
        Logger.log(ex + " ID: " + ID);
        attempts--;
        Utilities.sleep(2000);
      }
    }
    
    var initMap = {
      'Impressions': 0,'Conversions': 0,'Clicks': 0, 'CPC': 0,
      'Cost': 0, 'ConversionValue': 0, 'ROAS': 0
    };
    
    var accId = AdWordsApp.currentAccount().getCustomerId().replace(/-/g, '');  
    var monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                      'July', 'August', 'September', 'October', 'November', 'December'];
    
    var rows = results.getRows();
    for(var k in rows) {
      if(rows[k][3] != accId) { continue; }
      
      var monthKey = monthNames[parseInt(rows[k][1],10)] + ' ' + rows[k][2];
      if(!map[monthKey]) {
        map[monthKey] = {};
        map[monthKey][monthKey] = JSON.parse(JSON.stringify(initMap));
        map[monthKey]['Week 1'] = JSON.parse(JSON.stringify(initMap));
        map[monthKey]['Week 2'] = JSON.parse(JSON.stringify(initMap));
        map[monthKey]['Week 3'] = JSON.parse(JSON.stringify(initMap));
        map[monthKey]['Week 4'] = JSON.parse(JSON.stringify(initMap));
        map[monthKey]['Remaining Days'] = JSON.parse(JSON.stringify(initMap));
      }
      
      
      var date = parseInt(rows[k][0].substring(6,8),10);
      
      var key = 'Remaining Days';
      if(date < 8) { 
        key = 'Week 1'; 
      } else if(date < 15) { 
        key = 'Week 2'; 
      } else if(date < 22) { 
        key = 'Week 3'; 
      } else if(date < 29) { 
        key = 'Week 4'; 
      }
      
      map[monthKey][key].Impressions += parseInt(rows[k][5],10);
      map[monthKey][key].Clicks += parseInt(rows[k][6],10);
      //map[monthKey][key].Conversions += parseInt(rows[k][8].toString().replace(/,/g,''),10);
      map[monthKey][key].Cost += parseFloat(rows[k][7].toString().replace(/,/g,''));
      map[monthKey][key].ConversionValue += parseFloat(rows[k][4].toString().replace(/,/g,''));
      
      
      map[monthKey][monthKey].Cost += parseFloat(rows[k][7].toString().replace(/,/g,''));
      map[monthKey][monthKey].ConversionValue += parseFloat(rows[k][4].toString().replace(/,/g,''));
      map[monthKey][monthKey].Impressions += parseInt(rows[k][5],10);
      map[monthKey][monthKey].Clicks += parseInt(rows[k][6],10);
      //map[monthKey][monthKey].Conversions += parseInt(rows[k][8].toString().replace(/,/g,''),10);
    }
    
    return map;
  }
  
  function getStatsFromAdWords(startDate, endDate) {
    
    var initMap = {
      'Impressions': 0,'Conversions': 0,'Clicks': 0, 'CPC': 0,
      'Cost': 0, 'ConversionValue': 0, 'ROAS': 0
    };
    
    var DATE_RANGE = startDate.replace(/-/g,'') + ',' + endDate.replace(/-/g,'');
    var OPTIONS = { includeZeroImpressions : true };
    var cols = ['Date','MonthOfYear','Year','Impressions','Conversions','Clicks','Cost'];
    var report = 'ACCOUNT_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'during',DATE_RANGE].join(' ');
    
    var results = {};
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      
      var monthKey = [row.MonthOfYear, row.Year].join(' ');
      if(!results[monthKey]) {
        results[monthKey] = {};
        results[monthKey][monthKey] = JSON.parse(JSON.stringify(initMap));
        results[monthKey]['Week 1'] = JSON.parse(JSON.stringify(initMap));
        results[monthKey]['Week 2'] = JSON.parse(JSON.stringify(initMap));
        results[monthKey]['Week 3'] = JSON.parse(JSON.stringify(initMap));
        results[monthKey]['Week 4'] = JSON.parse(JSON.stringify(initMap));
        results[monthKey]['Remaining Days'] = JSON.parse(JSON.stringify(initMap));
      }
      
      var date = parseInt(row.Date.split('-')[2],10);
      
      var key = 'Remaining Days';
      if(date < 8) { 
        key = 'Week 1'; 
      } else if(date < 15) { 
        key = 'Week 2'; 
      } else if(date < 22) { 
        key = 'Week 3'; 
      } else if(date < 29) { 
        key = 'Week 4'; 
      }
      
      //results[monthKey][key].Impressions += parseInt(row.Impressions,10);
      //results[monthKey][key].Clicks += parseInt(row.Clicks,10);
      results[monthKey][key].Conversions += parseInt(row.Conversions.toString().replace(/,/g,''),10);
      //results[monthKey][key].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
      
      //results[monthKey][monthKey].Impressions += parseInt(row.Impressions,10);
      //results[monthKey][monthKey].Clicks += parseInt(row.Clicks,10);
      results[monthKey][monthKey].Conversions += parseInt(row.Conversions.toString().replace(/,/g,''),10);
      //results[monthKey][monthKey].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
    }
    
    return results;
  }
  
  function getAdWordsFormattedDate(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
  }
  
  function round(num,n) {    
    return +(Math.round(num + "e+"+n)  + "e-"+n);
  }
  
  function info(msg) {
    Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);
  }