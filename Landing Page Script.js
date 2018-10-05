var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1vLlnJ5774_utpsfWVXWHgjHXevZLhqfyc7vTucbMcH8/edit#gid=1625314011';

function main() {
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Management Dashboard');
  var values = tab.getDataRange().getValues();
  values.shift();
  values.shift();
  
  var names = {};
  for(var z in values) {
    if(!values[z][0]) { continue; }
    names[values[z][0]] = 1;
  }
  
  MccApp.accounts().withCondition('Name IN ["' + Object.keys(names).join('","') + '"]').executeInParallel('run');
  
}

function run(){
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Management Dashboard');
  var values = tab.getDataRange().getValues();
  values.shift();
  values.shift();
  
  var accName = AdWordsApp.currentAccount().getName();
  var labelSettings = {};
  for(var z in values) {
    if(values[z][0] != accName) { continue; }
    var pageStats = getStatsForLandingPage(values[z][2], values[z][4])
    if(!pageStats) { continue; }
    tab.getRange(parseInt(z,10)+3,6,1,pageStats.length).setValues([pageStats]);
  }
}

function getStatsForLandingPage(label, publishDate) {
  var agIds = [];
  var adGroups = AdWordsApp.adGroups()
  .withCondition('LabelNames CONTAINS_ANY ["' + label + '"]')
  .get();
  
  while(adGroups.hasNext()) {
    agIds.push(adGroups.next().getId());
  }
  
  if(!agIds.length) {
    return '';
  }
  
  publishDate.setHours(12);
  var publishDateString = Utilities.formatDate(publishDate, 'PST', 'yyyyMMdd');
  
  var dt = Utilities.formatDate(publishDate, 'PST', 'MM/dd/yyyy');
  
  var dateTemp = new Date(dt);
  dateTemp.setDate(dateTemp.getDate()-1);
  var endDate = Utilities.formatDate(dateTemp, 'PST', 'yyyyMMdd');
  
  dateTemp.setDate(dateTemp.getDate()-29);
  var startDate = Utilities.formatDate(dateTemp, 'PST', 'yyyyMMdd');
  
  var statsOld = getStatsForDates(startDate+','+endDate, agIds);
  
  var dateTemp = new Date(dt);
  dateTemp.setDate(dateTemp.getDate()+29);
  var endDate = Utilities.formatDate(dateTemp, 'PST', 'yyyyMMdd');
  
  var statsNew = getStatsForDates(publishDateString+','+endDate, agIds);
  
  var statsNow = getStatsForDates('LAST_30_DAYS', agIds);
  
  return [
    statsOld['Computers'].CR, statsNew['Computers'].CR, statsNow['Computers'].CR,
    statsOld['Computers'].Cost, statsNew['Computers'].Cost, statsNow['Computers'].Cost,
    statsOld['Mobile devices with full browsers'].CR, statsNew['Mobile devices with full browsers'].CR,
    statsNow['Mobile devices with full browsers'].CR, statsOld['Mobile devices with full browsers'].Cost,
    statsNew['Mobile devices with full browsers'].Cost, statsNow['Mobile devices with full browsers'].Cost
  ];
}

function getStatsForDates(DATE_RANGE, agIds) {
  //Logger.log(DATE_RANGE);
  var initMap = { 'Clicks': 0, 'Cost': 0, 'Conversions': 0 };
  
  var stats = {
    'Mobile devices with full browsers': JSON.parse(JSON.stringify(initMap)),
    'Computers': JSON.parse(JSON.stringify(initMap))
  }
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['AdGroupId', 'Device', 'Conversions', 'Clicks', 'Cost'];
  var report = 'ADGROUP_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where AdGroupId IN [' + agIds.join(',') + ']',
               'during',DATE_RANGE].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(!stats[row.Device]) { continue; }
    
    stats[row.Device].Clicks += parseInt(row.Clicks,10);
    stats[row.Device].Conversions += parseFloat(row.Conversions.toString().replace(/,/g,''));
    stats[row.Device].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
  }
  
  for(var device in stats) {
    stats[device].CR = stats[device].Clicks == 0 ? 0 : (100*stats[device].Conversions / stats[device].Clicks)+'%';
  }
  
  return stats;
}