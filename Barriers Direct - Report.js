var MERCHANT_ID = 2792358;
var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1yd_wxeiYsCGi7NyGMDieBW7j7UhGBm48ClB6UBQNRRY/edit#gid=0';

function main() {
  MccApp.accounts().withCondition('Name = "Barriers Direct"').executeInParallel('run');
}

function run() {
  var itemMap = {};
  
  var TO = getAdWordsFormattedDate(1, 'yyyyMMdd');
  var FROM = TO.substring(0,6) + '01';
  var DR = FROM + ',' + TO;
  
  var date = new Date(getAdWordsFormattedDate(1, 'MMM d, yyyy'));
  date.setHours(12);
  date.setMonth(date.getMonth()-1);
  var TO = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  date.setDate(1);
  var FROM = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  
  var DR_2 = FROM + ',' + TO;
  compileItemIdReportForDate('MTD vs Last MTD', DR, DR_2, itemMap);
  
  var date = new Date(getAdWordsFormattedDate(0, 'MMM d, yyyy'));
  date.setHours(12);
  date.setDate(1);
  date.setDate(0);
  date.setDate(1);
  date.setDate(0);
  
  var TO = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  date.setDate(1);
  var FROM = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  
  compileItemIdReportForDate('Last Month vs Prior Month', 'LAST_MONTH', FROM+','+TO, itemMap);
  
  var TO = getAdWordsFormattedDate(1, 'yyyyMMdd');
  var FROM = getAdWordsFormattedDate(30, 'yyyyMMdd');
  var DR = FROM + ',' + TO;
  
  var date = new Date(getAdWordsFormattedDate(1, 'MMM d, yyyy'));
  date.setHours(12);
  date.setYear(date.getYear()-1);
  var TO = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  date.setDate(date.getDate()-30);
  var FROM = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  
  var DR_2 = FROM + ',' + TO;
  compileItemIdReportForDate('Last 30 Days (YoY)', DR, DR_2, itemMap);
  
  compileNonConvertingItemsReport();
  compileNonConvertingQueryReport();
  
  manageBids();
  //excludeItems(itemMap);
}

function manageBids() {
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Bid Manager');
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var map = {};
  for(var z in data) {
    var key = [data[z][0], data[z][1], data[z][2], data[z][3]].join('::::');
    map[key] = {
      'index': z, 'row': data[z]
    }
  }
  
  var OPTIONS = { 'includeZeroImpressions' : true };
  var cols = ['CampaignName', 'AdGroupName', 'AdGroupId', 'Id', 'ProductGroup','CpcBid','Cost','Conversions'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               "and CampaignStatus = ENABLED",
               "and AdGroupStatus = ENABLED",
               'during','LAST_30_DAYS'].join(' ');
  
  var toExclude = [], newBids = {}, newBidIds = [];
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    var supplier = extractCustomLabel(row.ProductGroup);
    var key = [supplier, row.CampaignName, row.AdGroupName, item].join('::::');
    var productKey = [row.AdGroupId, row.Id].join('-');
    
    if(map[key]) {
      var input = map[key];
      data[input['index']][6] = row.CpcBid;
      data[input['index']][5] = row.Conversions;
      data[input['index']][4] = row.Cost;
      var bid = parseFloat(row.CpcBid);
      if(input['row'][7] !== '' && input['row'][7] != bid) {
        newBids[productKey] = input['row'][7]; 
        newBidIds.push([row.AdGroupId, row.Id]);
      } else if(row.CpcBid != 'Excluded' && input['row'][8] == 'Yes') {
        toExclude.push([row.AdGroupId, row.Id]);
      }
    } else {
      data.push([supplier, row.CampaignName, row.AdGroupName, item, row.Cost, row.Conversions, row.CpcBid, '', ''])
    }
  }
  
  tab.getRange(2,1,data.length,data[0].length).setValues(data);
  
  if(toExclude.length > 0) {
    var iter = AdWordsApp.productGroups().withIds(toExclude).get();
    while(iter.hasNext()) {
      iter.next().exclude(); 
    }
  }
  
  if(newBidIds.length > 0) {
    var iter = AdWordsApp.productGroups().withIds(newBidIds).get();
    while(iter.hasNext()) {
      var pg = iter.next();
      var key = [pg.getAdGroup().getId(), pg.getId()].join('-');
      if(newBids[key]) {
        pg.setMaxCpc(newBids[key]) 
      }
    }
  }
  
}

function excludeItems(itemMap) {
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Items to Exclude'); 
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var toExclude = [];
  for(var z in data) {
    if(!data[z][0]) { continue; }
    if(!itemMap[data[z][0]]) { continue; }
    
    toExclude = toExclude.concat(itemMap[data[z][0]]);
  }
  
  if(!toExclude.length) { return; }
  
  var iter = AdWordsApp.productGroups().withIds(toExclude).get();
  while(iter.hasNext()) {
   var pg = iter.next();
    pg.exclude();
  }
}

function compileNonConvertingQueryReport() {
  var rows = []
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','AdGroupName', 'Query', 'Cost', 'Clicks', 'Impressions', 'AverageCpc', 'Ctr'];
  var report = 'SEARCH_QUERY_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where Conversions < 1",
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    
    rows.push([row.CampaignName, row.AdGroupName, row.Query, row.Clicks, row.Impressions, 
               row.Ctr, row.Cost, row.AverageCpc]);
  }
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Non Converting Queries');
  tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  
  if(!rows.length) {
    return; 
  }
  
  tab.getRange(3,1,rows.length,rows[0].length).setValues(rows);
}


function compileNonConvertingItemsReport() {
  var rows = []
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','AdGroupName', 'Id', 'ProductGroup', 'Cost', 'Clicks', 'Impressions', 'AverageCpc', 'Ctr'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               'and Conversions < 1',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item/* || !map[item]*/) { continue; }
    
    rows.push([row.CampaignName, row.AdGroupName, item, row.Clicks, row.Impressions, row.Ctr, row.Cost, row.AverageCpc]);
  }
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Non Converting Items');
  tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  
  if(!rows.length) {
    return; 
  }
  
  tab.getRange(3,1,rows.length,rows[0].length).setValues(rows);
}

function compileItemIdReportForDate(TAB_NAME, DR, DR_2, itemMap) {
  Logger.log('Fetching Merchant Center Data');
  //var map = fetchMerchantCentreData();
  
  var initMap = {
    'Clicks': 0, 'Impressions': 0, 'ctr': 0, 'Cost': 0, 'cpc': 0, 'Conversions':0, 
    'cpa': 0, 'cr': 0, 'ConversionValue': 0, 'roas': 0
  };
  
  var stats = {};
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','Id', 'ProductGroup', 'Cost', 'ConversionValue', 
              'Clicks', 'Impressions', 'Conversions'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               'during',DR].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item/* || !map[item]*/) { continue; }
    
    if(!itemMap[item]) {
      itemMap[item] = { 
        'IDS': [],
        'KEYS': {},
        'Supplier': extractCustomLabel(row.ProductGroup),
        'ProductType':  extractProductType(row.ProductGroup)
      };
    }
    
    var key = [row.AdGroupId, row.Id].join('--');
    if(!itemMap[item]['KEYS'][key]) {
      itemMap[item]['KEYS'][key] = 1;
      itemMap[item]['IDS'].push([row.AdGroupId, row.Id])
    }
    
    
    if(!stats[item]) {
      stats[item] = {
        'LM': JSON.parse(JSON.stringify(initMap)),
        'PM': JSON.parse(JSON.stringify(initMap))
      };
    }
    
    stats[item]['LM'].Clicks += parseInt(row.Clicks, 10);
    stats[item]['LM'].Impressions += parseInt(row.Impressions, 10);
    stats[item]['LM'].Conversions += parseInt(row.Conversions, 10);
    
    stats[item]['LM'].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
    stats[item]['LM'].ConversionValue += parseFloat(row.ConversionValue.toString().replace(/,/g,''));
  }
  
  
  var date = new Date(getAdWordsFormattedDate(0, 'MMM d, yyyy'));
  date.setHours(12);
  date.setDate(1);
  date.setDate(0);
  date.setDate(1);
  date.setDate(0);
  
  var TO = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  date.setDate(1);
  var FROM = Utilities.formatDate(date, 'PST', 'yyyyMMdd');
  
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               'during',DR_2].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item/* || !map[item]*/) { continue; }
    
    if(!itemMap[item]) {
      itemMap[item] = { 
        'IDS': [],
        'KEYS': {},
        'Supplier': extractCustomLabel(row.ProductGroup),
        'ProductType':  extractProductType(row.ProductGroup)
      };
    }
    
    var key = [row.AdGroupId, row.Id].join('--');
    if(!itemMap[item]['KEYS'][key]) {
      itemMap[item]['KEYS'][key] = 1;
      itemMap[item]['IDS'].push([row.AdGroupId, row.Id])
    }
    
    if(!stats[item]) {
      stats[item] = {
        'LM': JSON.parse(JSON.stringify(initMap)),
        'PM': JSON.parse(JSON.stringify(initMap))
      };
    }
    
    stats[item]['PM'].Clicks += parseInt(row.Clicks, 10);
    stats[item]['PM'].Impressions += parseInt(row.Impressions, 10);
    stats[item]['PM'].Conversions += parseInt(row.Conversions, 10);
    
    stats[item]['PM'].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
    stats[item]['PM'].ConversionValue += parseFloat(row.ConversionValue.toString().replace(/,/g,''));    
  }
  
  
  var out = [];
  for(var item in stats) {
    var map = itemMap[item];
    var row = [map.Supplier, map.ProductType, item];
    
    stats[item]['LM'].ctr = stats[item]['LM'].Impressions == 0 ? 0 : round(stats[item]['LM'].Clicks/stats[item]['LM'].Impressions, 4);
    stats[item]['LM'].cr = stats[item]['LM'].Clicks == 0 ? 0 : round(stats[item]['LM'].Conversions/stats[item]['LM'].Clicks, 4);
    stats[item]['LM'].cpc = stats[item]['LM'].Clicks == 0 ? 0 : round(stats[item]['LM'].Cost/stats[item]['LM'].Clicks, 2);
    stats[item]['LM'].cpa = stats[item]['LM'].Conversions == 0 ? 0 : round(stats[item]['LM'].Cost/stats[item]['LM'].Conversions, 2);    
    stats[item]['LM'].roas = stats[item]['LM'].Cost == 0 ? 0 : round(stats[item]['LM'].ConversionValue/stats[item]['LM'].Cost, 4);
    
    stats[item]['PM'].ctr = stats[item]['PM'].Impressions == 0 ? 0 : round(stats[item]['PM'].Clicks/stats[item]['PM'].Impressions, 4);
    stats[item]['PM'].cr = stats[item]['PM'].Clicks == 0 ? 0 : round(stats[item]['PM'].Conversions/stats[item]['PM'].Clicks, 4);
    stats[item]['PM'].cpc = stats[item]['PM'].Clicks == 0 ? 0 : round(stats[item]['PM'].Cost/stats[item]['PM'].Clicks, 2);
    stats[item]['PM'].cpa = stats[item]['PM'].Conversions == 0 ? 0 : round(stats[item]['PM'].Cost/stats[item]['PM'].Conversions, 2);    
    stats[item]['PM'].roas = stats[item]['PM'].Cost == 0 ? 0 : round(stats[item]['PM'].ConversionValue/stats[item]['PM'].Cost, 4);
    
    for(var metric in stats[item]['LM']) {
      row.push(stats[item]['LM'][metric], stats[item]['PM'][metric], 
               stats[item]['PM'][metric] == 0 ? 'N/A' : 100*((stats[item]['LM'][metric] - stats[item]['PM'][metric]) / stats[item]['PM'][metric])+'%');
    }
    
    out.push(row);
  }
  
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName(TAB_NAME);
  tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  tab.getRange(3,1,out.length,out[0].length).setValues(out);
  tab.sort(3);
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}


/*

function fetchMerchantCentreData() {
  var map = {};
  
  var pageToken;
  var pageNum = 1;
  var maxResults = 250;
  
  try {
    // List all the products for a given merchant.
    do {
      var products = ShoppingContent.Products.list(MERCHANT_ID, {
        pageToken: pageToken,
        maxResults: maxResults
      });
      
      if (products.resources) {
        for(var z in products.resources) {
          var offerId = products.resources[z].offerId.replace(/ /g,'').toLowerCase();
          map[offerId] = { 'title': products.resources[z].title, 'price': products.resources[z].price.value };
        }
      } 
      
      pageToken = products.nextPageToken;
      pageNum++;
    } while (pageToken);
  } catch(ex){
    Logger.log(MERCHANT_ID + ': ' + ex);
    throw ex;
  }
  
  //throw eee;
  return map;
}
*/

function extractItemId(str) {
  //return str.match(/(?:"[^"]*"|^[^"]*$)/)[0].replace(/"/g, "");
  var matches = str.match(/item id = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}


function extractProductType(str) {
  var matches = str.match(/product type = \"(.*?)\"/g);
  if(matches) { 
    if(matches[1]) {
      return  matches[1].replace('product type = ', '').replace(/\"/g,'')
    }
    
    return matches[0].replace('product type = ', '').replace(/\"/g,'')  
  }
  
  return ''
}

function extractCustomLabel(str) {
  var matches = str.match(/custom label 0 = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}
