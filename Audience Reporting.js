var url = 'https://docs.google.com/spreadsheets/d/1CsdkuctJ7CoFm2WxSRnRfQF66eVCojZypsKj2ZJXK_c/edit'; 

function main() {
  MccApp.accounts().withCondition('Name = "Golfbreaks"').executeInParallel('run');
}

function getIdToName() {
  var map = {};
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Audience Mapping');
  if(!tab) { return {}; }
  
  var data = tab.getDataRange().getValues();
  var header = data.shift();
  
  var index = header.indexOf('User list name');
  var index_2 = header.indexOf('Criterion ID');
  			
  for(var z in data) {
    map[data[z][index_2]] = data[z][index]; 
  }
  
  return map;
}

function run() {
  var map = getIdToName();
  
  var initMap = {
    'Impressions': 0,'Clicks': 0,'Ctr': 0,'Cost': 0,'Conversions': 0,
    'CostPerConversion': 0,'AverageCpc': 0,'AveragePosition': 0,'ConversionRate': 0
  };
  
  var out = [['List Name', 'Impressions','Clicks','Ctr','Cost','Conversions',
              'CPA','Average Cpc','Average Position','Conversion Rate']];
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['Criteria','Id','UserListName','Impressions','Clicks','Ctr',
              'Cost','Conversions','CostPerConversion','AverageCpc','AveragePosition','ConversionRate'];
  
  var reportName = 'AUDIENCE_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'during','LAST_30_DAYS'].join(' ');
  
  var statsMap = {};             
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var name = map[row.Id] ? map[row.Id] : row.UserListName;
    if(!name) { continue; }
    if(!statsMap[name]) {
      statsMap[name] = JSON.parse(JSON.stringify(initMap));
    }
    
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.AveragePosition = parseFloat(row.AveragePosition);
    row.Impressions = parseInt(row.Impressions,10);
    row.Clicks = parseInt(row.Clicks,10);
    row.Conversions = parseInt(row.Conversions,10);
    
    statsMap[name].Clicks += row.Clicks;
    statsMap[name].Impressions += row.Impressions;
    statsMap[name].Conversions += row.Conversions;
    statsMap[name].Cost += row.Cost;
    statsMap[name].AveragePosition += row.AveragePosition*row.Impressions;		
  }
  
  for(var name in statsMap) {
    statsMap[name].AverageCpc =  statsMap[name].Clicks == 0 ? 0 : round(statsMap[name].Cost / statsMap[name].Clicks, 2);
    statsMap[name].AveragePosition =  statsMap[name].Impressions == 0 ? 0 : round(statsMap[name].AveragePosition / statsMap[name].Impressions, 1);    
    statsMap[name].CostPerConversion =  statsMap[name].Conversions == 0 ? 0 : round(statsMap[name].Cost / statsMap[name].Conversions, 2);    
    statsMap[name].Ctr = round(100*(statsMap[name].Clicks / statsMap[name].Impressions), 2)+'%';
    statsMap[name].ConversionRate = statsMap[name].Clicks == 0 ? 0 : round(100*(statsMap[name].Conversions / statsMap[name].Clicks), 2)+'%';    
  }
  
  for(var name in statsMap) {
    var row = [name];
    for(var metric in statsMap[name]) {
      row.push(statsMap[name][metric]);
    }
    
    out.push(row);
  }
  
  var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Report');
  tab.clearContents();
  tab.getRange(1,1,out.length,out[0].length).setValues(out).setFontFamily('Calibri');
  tab.getRange(1,1,1,out[0].length).setBackground('#efefef').setFontWeight('bold');
  tab.setFrozenRows(1);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}