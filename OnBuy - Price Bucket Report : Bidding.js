var ID = 98079839, MERCHANT_ID = 112872187;
var URL = 'https://docs.google.com/spreadsheets/d/1DDjTa_y0N5_HIutTiyjEbceOltuKabUsiB-BlB9nAQI/edit';
var FULL_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1lCv9JPk3UNdJdaS1s11KsddChlu5Zcr82Eq4lavlNtM/edit';
var EMAIL = 'cas.paton@onbuy.com';
var CC = ['neeraj@pushgroup.co.uk','charlie@pushgroup.co.uk','ricky@pushgroup.co.uk',
          'ian@pushgroup.co.uk','sandeep@pushgroup.co.uk','jay@pushgroup.co.uk'];

function main() {
  MccApp.accounts().withCondition('Name = "OnBuy"').executeInParallel('run');
}

function run() {
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'), 10);
  
  if(hour == 6) {
    compilePriceBucketReport();
  }
  
  if(hour == 12) {
    compileAdGroupCampaignReport(); 
  }
  
  updateBids();
}

function compileAdGroupCampaignReport() {
  var output = [];
  var cols = ['CampaignName', 'Cost', 'Clicks', 'Conversions', 'CostPerConversion', 'ConversionValue'];
  var query = 'SELECT ' + cols.join(',') + ' FROM CAMPAIGN_PERFORMANCE_REPORT DURING LAST_30_DAYS';
  var rows = AdWordsApp.report(query, { 'includeZeroImpressions': false } ).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    var out = [];
    for(var i in cols) {
      out.push(row[cols[i]]);
    }
    
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    out.push(row.Cost == 0 ? 0 : row.ConversionValue/row.Cost);
    
    output.push(out);
  }
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Category Report');
  tab.getRange(2, 1, tab.getLastRow(), tab.getLastColumn()).clearContent();
  tab.getRange(2, 1, output.length, output[0].length).setValues(output);
  tab.sort(6, false);
  
  var map = {}
  var cols = ['CampaignName', 'AdGroupName', 'Cost', 'Clicks', 'Conversions', 'CostPerConversion', 'ConversionValue'];
  var query = 'SELECT ' + cols.join(',') + ' FROM ADGROUP_PERFORMANCE_REPORT DURING LAST_30_DAYS';
  var rows = AdWordsApp.report(query, { 'includeZeroImpressions': false } ).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    if(!map[row.AdGroupName]) {
      map[row.AdGroupName] = { 'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0 };
    }
    
    
    map[row.AdGroupName].Clicks += parseInt(row.Clicks, 10);
    map[row.AdGroupName].Conversions += parseFloat(row.Conversions.toString().replace(/,/g,''));
    map[row.AdGroupName].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
    map[row.AdGroupName].ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
  }
  
  var output = [];
  for(var key in map) {
    var row = map[key];
    row.CPA = row.Conversions == 0 ? 0 : row.Cost/row.Conversions;
    row.ROAS = row.Cost == 0 ? 0 : row.ConversionValue/row.Cost;
    output.push([key, row.Cost, row.Clicks, row.Conversions, row.CPA, row.ConversionValue, row.ROAS]);
  }
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Sub Category Report');
  tab.getRange(2, 1, tab.getLastRow(), tab.getLastColumn()).clearContent();
  tab.getRange(2, 1, output.length, output[0].length).setValues(output);
  tab.sort(6, false);  
}

function updateBids() {
  
  var bid = {};
  var ss = SpreadsheetApp.openByUrl(URL);
  var tabs = ss.getSheets();
  for(var k in tabs) {
    var camp = tabs[k].getName();
    
    if(camp.indexOf('Price Bucket') > -1) { continue; }
    if(camp.indexOf('%') < 0) { continue; }       
    
    var data = tabs[k].getDataRange().getValues();
    data.shift();
    data.shift();
    
    bid[camp] = {};
    
    for(var z in data) {
      if(!data[z][7]) { continue; }
      bid[camp][data[z][0]] = data[z][7];
    }
    
    if(!Object.keys(bid[camp]).length) {
      delete bid[camp];
    }
  }
  
  var names = Object.keys(bid);
  if(!names.length) { return; }
  
  var ids = [], bidMap = {};
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['CampaignName', 'AdGroupId', 'Id', 'ProductGroup', 'CpcBid'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where ProductGroup CONTAINS "custom label 1"',
               'and CampaignStatus = ENABLED',
               'and AdGroupStatus = ENABLED',
               'and ProductGroup CONTAINS "-"',
               'and CampaignName IN ["' + names.join('","') + '"]',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractCustomLabel(row.ProductGroup);
    if(!bid[row.CampaignName][item]) { continue; }
    
    row.CpcBid = parseFloat(row.CpcBid);
    if(bid[item] == row.CpcBid) { continue; }
    
    ids.push([row.AdGroupId, row.Id]);
    bidMap[[row.AdGroupId, row.Id].join('-')] = bid[row.CampaignName][item];
  }
  
  var chunk_size = 9500;
  var groups = ids.map( function(e,i){ 
    return i%chunk_size===0 ? ids.slice(i,i+chunk_size) : null; 
  }).filter(function(e){ return e; });
  
  
  for(var z in groups) { 
    var iter = AdWordsApp.productGroups()
    .withIds(groups[z]).get();
    
    while(iter.hasNext()) {
      var pg = iter.next();
      var newBid = bidMap[[pg.getAdGroup().getId(), pg.getId()].join('-')];
      pg.setMaxCpc(newBid);
    }	
  }
  
}

function compilePriceBucketReport() {
  var stats = {
    'Price Bucket - 6%': {},
    'Price Bucket - 9%': {}
  };
  
  Logger.log('Fetching Price Bucket Stats');  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','Id', 'ProductGroup', 'Cost','ConversionValue','Clicks','Conversions'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where ProductGroup CONTAINS "custom label 1"',
               'and CampaignName CONTAINS "%"',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractCustomLabel(row.ProductGroup);
    if(!item) { continue; }
    
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    row.Clicks = parseInt(row.Clicks, 10);
    row.Conversions = parseInt(row.Conversions, 10);
    
    if(row.CampaignName.indexOf('6%') > -1) {
      if(!stats['Price Bucket - 6%'][item]) {
        stats['Price Bucket - 6%'][item] = { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0 } 
      }
      
      stats['Price Bucket - 6%'][item]['Cost'] += row.Cost;
      stats['Price Bucket - 6%'][item]['Conversions'] += row.Conversions;
      stats['Price Bucket - 6%'][item]['ConversionValue'] += row.ConversionValue;
      stats['Price Bucket - 6%'][item]['Clicks'] += row.Clicks;
      
    } else if(row.CampaignName.indexOf('9%') > -1) {
      if(!stats['Price Bucket - 9%'][item]) {
        stats['Price Bucket - 9%'][item] = { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0 } 
      }
      
      stats['Price Bucket - 9%'][item]['Cost'] += row.Cost;
      stats['Price Bucket - 9%'][item]['Conversions'] += row.Conversions;
      stats['Price Bucket - 9%'][item]['ConversionValue'] += row.ConversionValue;
      stats['Price Bucket - 9%'][item]['Clicks'] += row.Clicks;
    }
    
    
    if(!stats[row.CampaignName]) {
      stats[row.CampaignName] = {};
    }
    
    if(!stats[row.CampaignName][item]) {
      stats[row.CampaignName][item] = { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0 } 
    }
    
    stats[row.CampaignName][item]['Cost'] += row.Cost;
    stats[row.CampaignName][item]['Conversions'] += row.Conversions;
    stats[row.CampaignName][item]['ConversionValue'] += row.ConversionValue;
    stats[row.CampaignName][item]['Clicks'] += row.Clicks;
    
  }
  
  var ss = SpreadsheetApp.openByUrl(URL);
  for(var key in stats) {
    var tab = ss.getSheetByName(key);
    if(!tab) {
      ss.setActiveSheet(ss.getSheetByName('Price Bucket - 6%'));
      tab = ss.duplicateActiveSheet();
      tab.setName(key);
    }
    
    /*if(key != 'Price Bucket - 9%') {
    tab.getRange(3, 1, tab.getLastRow(), tab.getLastColumn()).clearContent();
    tab.getRange(row, column)
    }*/
    
    var toDelete = [];
    var data = tab.getDataRange().getValues();
    data.shift();
    data.shift();
    
    for(var z in data) {
      var item = data[z][0];
      if(!stats[key][item] || !item) { 
        toDelete.push(parseInt(z,10)+3);
        continue; 
      }
      
      var row = [stats[key][item].Cost, stats[key][item].Clicks,
                 stats[key][item].Conversions, stats[key][item].ConversionValue, 
                 stats[key][item].ConversionValue == 0 ? 0 :stats[key][item].Cost/stats[key][item].ConversionValue];
      tab.getRange(parseInt(z,10)+3, 2, 1, row.length).setValues([row]);
      delete stats[key][item];
    }
    
    for(var z = toDelete.length-1; z >= 0; z--) {
      tab.deleteRow(toDelete[z]);
    }
    
    var out = [];
    for(var item in stats[key]) {
      out.push([item, stats[key][item].Cost, stats[key][item].Clicks,
                stats[key][item].Conversions, stats[key][item].ConversionValue, 
                stats[key][item].ConversionValue == 0 ? 0 : stats[key][item].Cost/stats[key][item].ConversionValue]);
    }
    
    if(out.length) {
      tab.getRange(tab.getLastRow()+1, 1, out.length, out[0].length).setValues(out);
    }
  }
}

function extractCustomLabel(str) {
  var matches = str.match(/custom label 1 = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}