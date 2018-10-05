var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1QyNsJI2d_0UMo6yrW6M13Zjp4MUkGBCrwc2hgWNnV3o/edit';

function main() {
  MccApp.accounts().withIds(['368-834-9251']).executeInParallel('run');
}

function run() {
  var data = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(AdWordsApp.currentAccount().getName()).getDataRange().getValues();
  data.shift();
  data.shift();
  
  
  var map = {};
  for(var z in data) {
    var key = [data[z][0],data[z][1]].join('!~!');
    map[key] = {
      'HL1a': data[z][3],	
      'HL2a': data[z][4],
      'DLa': data[z][5],
      'HL1b': data[z][6],	
      'HL2b': data[z][7],
      'DLb': data[z][8],
      'HL1c': data[z][9],
      'HL2c': data[z][10],
      'DLc': data[z][11]
    };
  }
  
  var columnHeads = [
    'Campaign', 'Ad group', 'Headline 1', 'Headline 2', 'Description',
    'Path 1', 'Path 2', 'Final URL', 'Ad state'
  ];
  
  var upload = AdWordsApp.bulkUploads().newCsvUpload(columnHeads, {'moneyInMicros': false});

  var query = [
    'SELECT CampaignName, AdGroupName, HeadlinePart1, HeadlinePart2, Description, Path1, Path2, CreativeFinalUrls',
    'FROM AD_PERFORMANCE_REPORT',
    'WHERE CampaignStatus = ENABLED and AdGroupStatus = ENABLED and Status = ENABLED',
    'and HeadlinePart1 CONTAINS_IGNORE_CASE "PushAds"',
    'DURING TODAY'
  ].join(' ');
  
  var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
  while(rows.hasNext()) {
    var row = rows.next(); 
    var key = [row.CampaignName, row.AdGroupName].join('!~!');
    if(!map[key]) { continue; }
    
    row.Url = JSON.parse(row.CreativeFinalUrls)[0];
    
    var type = '';
    if(row.HeadlinePart1.indexOf('HL1a') > -1) {
     type = 'a' 
    } else if(row.HeadlinePart1.indexOf('HL1b') > -1) {
     type = 'b' 
    } else if(row.HeadlinePart1.indexOf('HL1c') > -1) {
     type = 'c' 
    }
    
    if(!type) { continue; }
    
    var headline1 = map[key]['HL1'+type],
        headline2 = map[key]['HL2'+type],
        desc = map[key]['DL'+type];
    
    upload.append({
      'Campaign': row.CampaignName, 
      'Ad group': row.AdGroupName, 
      'Headline 1': headline1,
      'Headline 2': headline2, 
      'Description': desc,
      'Path 1': row.Path1, 
      'Path 2': row.Path2, 
      'Final URL': row.Url, 
      'Ad state': 'enabled'
    });
  }
  
  Logger.log('applying');
  upload.apply();
}