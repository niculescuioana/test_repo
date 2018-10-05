var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1nSvnNWBhz_oERnzWJ9xTGKA5F6LTeHapNskPhAcCrZQ/edit#gid=549598045';

function main() {
  MccApp.accounts().withIds(['387-910-8272']).executeInParallel('run');
}

function run() {
  var statsByLabel = {
    'DK': {
      'ID': 2728236,
      'TabName': 'DK - Denmark',
      'SearchImpressions': 0,
      'Clicks': 0, 'Impressions': 0, 'CTR': 0, 'AvgCPC': 0,	'Cost': 0,
      'AvgPosition': 0, 'Conversions': 0, 'CPA': 0,	'CR': 0, 'AllConversions': 0, 
      'DK Transactions (www.scanteak.dk)': 0, 'DK Order confirmation': 0, 'Scanteak Calls From ads': 0
    },
    'SE': {
      'ID': 123148380,
      'TabName': 'SE - Sweden',
      'SearchImpressions': 0,
      'Clicks': 0, 'Impressions': 0, 'CTR': 0, 'AvgCPC': 0,	'Cost': 0,
      'AvgPosition': 0, 'Conversions': 0, 'CPA': 0,	'CR': 0, 'AllConversions': 0, 
      'SE Transactions (www.scanteak.se)': 0, 'SE Order confirmation': 0, 'Scanteak Calls From ads': 0
    }
  }
  
  
  var month = getAdWordsFormattedDate(1, 'MMMM yyyy');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  
  var query = [
    'SELECT CampaignId, Labels, Clicks, Impressions, Cost, Conversions, AllConversions, AveragePosition',
    'FROM CAMPAIGN_PERFORMANCE_REPORT DURING', FROM.replace(/-/g, '') + ',' + TO.replace(/-/g,'')
  ].join(' ');
  
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    var label = '';
    if(row['Labels'].indexOf('DK') > -1) {
      label = 'DK'; 
    } else if(row['Labels'].indexOf('SE') > -1) {
      label = 'SE'; 
    }
    
    if(!label) { continue; }
    
    
    row.Clicks = parseInt(row.Clicks, 10);
    row.Impressions = parseInt(row.Impressions, 10);
    
    row.AveragePosition = parseFloat(row.AveragePosition);
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g, ''));
    row.Conversions = parseFloat(row.Conversions.toString().replace(/,/g, ''));
    row.AllConversions = parseFloat(row.AllConversions.toString().replace(/,/g, ''));
    
    if(row.AveragePosition > 0) {
      statsByLabel[label]['SearchImpressions'] += row.Impressions;
      statsByLabel[label]['AvgPosition'] += row.AveragePosition*row.Impressions;
    }
    
    statsByLabel[label]['Clicks'] += row.Clicks;
    statsByLabel[label]['Cost'] += row.Cost;
    statsByLabel[label]['Impressions'] += row.Impressions;
    
    statsByLabel[label]['Conversions'] += row.Conversions;
    statsByLabel[label]['AllConversions'] += row.AllConversions;
  }
  
  var query = [
    'SELECT CampaignId, Labels, ConversionTypeName, Conversions',
    'FROM CAMPAIGN_PERFORMANCE_REPORT DURING', FROM.replace(/-/g, '') + ',' + TO.replace(/-/g,'')
  ].join(' ');
  
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    var label = '';
    if(row['Labels'].indexOf('DK') > -1) {
      label = 'DK'; 
    } else if(row['Labels'].indexOf('SE') > -1) {
      label = 'SE'; 
    }
    
    if(!label) { continue; }
    
    if(statsByLabel[label][row.ConversionTypeName] == undefined) {
      continue;
    }
    
    statsByLabel[label][row.ConversionTypeName] += parseFloat(row.Conversions.toString().replace(/,/g, ''));
  }
  
  var rowNum = '';
  for(var label in statsByLabel) {
    var row = statsByLabel[label];
    var tabName = row['TabName'];
    delete row['TabName'];
    
    row.AvgPosition = row.SearchImpressions == 0 ? 0 : row.AvgPosition / row.SearchImpressions;
    row.CTR = row.Impressions == 0 ? 0 : row.Clicks / row.Impressions;
    row.AvgCPC = row.Clicks == 0 ? 0 : row.Cost / row.Clicks;
    row.CR = row.Clicks == 0 ? 0 : row.Conversions / row.Clicks;
    row.CPA = row.Conversions == 0 ? 0 : row.Cost / row.Conversions;
    
    delete row['SearchImpressions'];
    
    var optArgs = { 'filters': 'ga:medium==cpc;ga:source==google' };
    var googleStats = { 'Sessions': 0, 'ConversionValue': 0 };
    getDataFromAnalytics(row['ID'],googleStats,FROM,TO,optArgs);
    
    var overallStats = { 'Sessions': 0, 'ConversionValue': 0 };
    var optArgs = {};
    getDataFromAnalytics(row['ID'],overallStats,FROM,TO,optArgs);
    
    row['TotalSessions'] = overallStats['Sessions'];
    row['PPCSessions'] = googleStats['Sessions'];
    row['PCTPPCSessions'] = row.TotalSessions == 0 ? 0 : round(row.PPCSessions / row.TotalSessions, 4)
    
    row['TotalRevenue'] = overallStats['ConversionValue'];
    row['PPCRevenue'] = googleStats['ConversionValue'];
    row['PCTPPCRevenue'] = row.TotalRevenue == 0 ? 0 : round(row.PPCRevenue / row.TotalRevenue, 4)
    
    row['PPCROAS'] = row['Cost'] == 0 ? 0 : row['PPCRevenue'] / row['Cost'];
    row['ROAS'] = row['Cost'] == 0 ? 0 : row['TotalRevenue'] / row['Cost'];
    
    delete row['ID'];
    
    var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(tabName);
    var out = [month];
    for(var key in row) {
      out.push(row[key]); 
    }
    
    tab.getRange('A:A').setNumberFormat('@STRING@');
    var data = tab.getDataRange().getValues();
    data.shift();
    data.shift();
    data.shift();
    data.shift();
    
    
    var found = false;
    for(var z in data) {
      if(data[z][0] == month) {
        rowNum = parseInt(z,10)+5;
        tab.getRange(parseInt(z,10)+5,1,1,out.length).setValues([out]);
        found = true;
        break; 
      }
    }
    
    if(!found) {
      tab.insertRowBefore(5); 
      rowNum = 5;
      tab.getRange(5,1,1,out.length).setValues([out]);
    }
  }
  
  if(rowNum) {
    var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('All');
    if(tab.getRange(rowNum,1).getValue() != month) {
      tab.insertRowBefore(5); 
    }
    tab.getRange(4,1,1,tab.getLastColumn()).copyTo(tab.getRange(rowNum, 1, 1, tab.getLastColumn()));
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



function getDataFromAnalytics(PROFILE_ID,stats,FROM,TO,optArgs) {
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:transactionRevenue,ga:sessions",
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
    stats.ConversionValue += parseFloat(rows[k][0]);
    stats.Sessions += parseInt(rows[k][1],10);  
  }
}