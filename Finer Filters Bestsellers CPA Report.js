function main() {
    var ids = ['603-721-5757'];
    MccApp.accounts().withIds(ids).executeInParallel('run');
  }
  
  function run() {
    var CAMPAIGN_NAME = "GS - Bestsellers";
    var URL = 'https://docs.google.com/spreadsheets/d/1G2JsbVCH8y1XbEiWPsYIYLWucuITHxn0ZI4V0WMBgjo/edit';
    
    var DATE_RANGE = 'LAST_30_DAYS';
    
    var days = {
      'Sunday': 1, 'Monday': 1, 'Tuesday': 1, 'Wednesday': 1, 'Thursday': 1, 'Friday': 1, 'Saturday': 1
    }
    
    var devices = {
      'Computers': 1, 'Mobile devices with full browsers': 1, 'Tablets with full browsers': 1
    }
    
    var stats = {}, pgMap = {};
    var OPTIONS = { 'includeZeroImpressions' : false };
    var cols = ['AdGroupId', 'Id', 'ProductGroup'];
    var report = 'PRODUCT_PARTITION_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'where ProductGroup CONTAINS_IGNORE_CASE "item"',
                 'and CampaignName = "' + CAMPAIGN_NAME + '"',               
                 'during',DATE_RANGE].join(' ');
    
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      if(row.ProductGroup.indexOf('id = *') > -1) { continue; }
      
      var pgId = row.ProductGroup.replace('* / item id = "', '').replace('"','')    
      pgMap[row.AdGroupId] = pgId;
    }
    
    var agMap = {};
    var cols = ['AdGroupId', 'HourOfDay', 'DayOfWeek', 'Device', 'Cost', 'Conversions'];
    var report = 'ADGROUP_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'where CampaignName = "' + CAMPAIGN_NAME + '"',
                 'during',DATE_RANGE].join(' ');
    
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      var pgId = pgMap[row.AdGroupId];
      agMap[row.AdGroupId] = pgId;
      if(!stats[pgId]) {
        stats[pgId] = {
          'hourly': {},
          'dow': {},
          'device': {}
        };
      }
      
      row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
      row.Conversions = parseFloat(row.Conversions.toString().replace(/,/g,''));
      
      if(!stats[pgId]['hourly'][row.HourOfDay]) {
        stats[pgId]['hourly'][row.HourOfDay] = { 'Cost': 0, 'Conversions': 0 } 
      }
      
      stats[pgId]['hourly'][row.HourOfDay].Cost += row.Cost;
      stats[pgId]['hourly'][row.HourOfDay].Conversions += row.Conversions;
      
      if(!stats[pgId]['dow'][row.DayOfWeek]) {
        stats[pgId]['dow'][row.DayOfWeek] = { 'Cost': 0, 'Conversions': 0 } 
      }
      
      stats[pgId]['dow'][row.DayOfWeek].Cost += row.Cost;
      stats[pgId]['dow'][row.DayOfWeek].Conversions += row.Conversions;
      
      if(!stats[pgId]['device'][row.Device]) {
        stats[pgId]['device'][row.Device] = { 'Cost': 0, 'Conversions': 0 } 
      }
      
      stats[pgId]['device'][row.Device].Cost += row.Cost;
      stats[pgId]['device'][row.Device].Conversions += row.Conversions;
    }
    
    var topKeywordsMap = getTopKeywords(DATE_RANGE, agMap);
    
    var ss = SpreadsheetApp.openByUrl(URL);
    var temp = ss.getSheetByName('Template');
    
    for(var id in stats) {
      var tab = ss.getSheetByName('ID ' + id);
      if(!tab) {
        ss.setActiveSheet(temp);
        tab = ss.duplicateActiveSheet(); 
        tab.setName('ID ' + id);
      }
      
      tab.showSheet();
      
      var map = stats[id];
      
      var out = [];
      for(var x=0; x<24; x++) {
        var row = map['hourly'][x];
        if(!row) {
          row = { 'Cost': 0, 'Conversions': 0 }
        }
        
        row.CPA = row.Conversions == 0 ? row.Cost : round(row.Cost/row.Conversions,2);
        
        out.push([x, row.CPA, row.Conversions]);
      }
      
      tab.getRange(3,1,out.length,3).setValues(out);
      
      
      out = [];
      for(var day in days) {
        var row = map['dow'][day];
        if(!row) {
          row = { 'Cost': 0, 'Conversions': 0 }
        }
        
        row.CPA = row.Conversions == 0 ? row.Cost : round(row.Cost/row.Conversions,2);
        
        out.push([day, row.CPA, row.Conversions]);
      }
      
      tab.getRange(3,5,out.length,3).setValues(out);
      
      out = [];
      for(var device in devices) {
        var row = map['device'][device];
        if(!row) {
          row = { 'Cost': 0, 'Conversions': 0 }
        }
        
        row.CPA = row.Conversions == 0 ? row.Cost : round(row.Cost/row.Conversions,2);
        
        out.push([device, row.CPA, row.Conversions]);
      }
      
      tab.getRange(3,9,out.length,3).setValues(out);
      
      tab.getRange(3, 13, tab.getLastRow()-2, 4).clearContent();
      
      if(topKeywordsMap[id] && topKeywordsMap[id].length) {
        tab.getRange(3, 13, topKeywordsMap[id].length, 4).setValues(topKeywordsMap[id]);
      }
    }
    
    /*var sheets = ss.getSheets();
    for(var x in sheets) {
     var sheet = sheets[x];
      if(sheet.getName() == 'Template') { continue; }
      
      var id = sheet.getName().replace('ID ', '')
      if(!stats[id]) { ss.deleteSheet(sheet); }
    }*/
  }
  
  
  function getTopKeywords(DATE_RANGE, agMap) {
    var agIds = Object.keys(agMap);
    
    var OPTIONS = { 'includeZeroImpressions' : false };
    var cols = ['AdGroupId', 'Query', 'Cost', 'Conversions'];
    var report = 'SEARCH_QUERY_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'where AdGroupId IN [' + agIds.join(',') + ']',
                 'and Cost > 0',
                 'during',DATE_RANGE].join(' ');
    
    var map = {};
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    while(reportIter.hasNext()) {
      var row = reportIter.next();
      var pgId = agMap[row.AdGroupId];
      if(!map[pgId]) {
        map[pgId] = [];
      }
      
      row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
      row.Conversions = parseFloat(row.Conversions.toString().replace(/,/g,''));
      row.CPA = row.Conversions == 0 ? row.Cost : round(row.Cost/row.Conversions,2);
      
      map[pgId].push([row.Query, row.Cost, row.Conversions, row.CPA]);
    }
    
    for(var id in map) {
      map[id].sort(function(a,b){return b[1]-a[1]; });
      if(map[id].length > 20) {
        map[id] = map[id].slice(0,21); 
      }
    }
    
    return map;
  }
  
  function round(num,n) {    
    return +(Math.round(num + "e+"+n)  + "e-"+n);
  }
  
  function getAdWordsFormattedDate(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
  }