var URL = 'https://docs.google.com/spreadsheets/d/1eNUCWI152ICqNitzqmAwyf_BzGd3vdpTS3cp8SP0QeI/edit#gid=1146861326';

function main() {
  var accounts = [
    'Adel Direct',
    'Fits My Samsung'
    //'Essentials London', 
    //'Barriers Direct'//, 
    //'OnBuy'
  ];
  
  MccApp.accounts()
  //.withCondition('Name = "OnBuy"')
  .withCondition('Name IN ["' + accounts.join('","') + '"]')
  .executeInParallel('run') ;
}

function run() {
  var sheet = SpreadsheetApp.openByUrl(URL).getSheetByName(AdWordsApp.currentAccount().getName())
  
  var inputs = sheet.getDataRange().getValues();
  var flag = inputs[0][1];
  
  if(flag == 'Pause Script') { return; }
  
  if(flag == 'Upload Data') {
    uploadNewKeywords(sheet, inputs); 
    Utilities.sleep(10000);
  }
  
  Logger.log('Reading Products from Merchant Center');
  var map = getApprovedShoppingProducts(inputs[3][1]);
  uploadParams(map, inputs);
  
  if(flag == 'Export Data') {
    exportConvertingShoppingQueries(map, sheet, inputs); 
  }
}

function uploadNewKeywords(sheet, inputs) {
  
  var campaignName = inputs[4][1];
  if(!campaignName || !AdWordsApp.campaigns().withCondition('Name = "' + campaignName + '"').get().next()) {
    return; 
  }
  
  var flag = false;
  var header = ['Campaign', 'Ad group', 'Ad group state', 'Default max. CPC'];
  var agUpload = AdWordsApp.bulkUploads().newCsvUpload(header,{'moneyInMicros': false});
  
  var header = ['Campaign', 'Ad group', 'Keyword', 'Keyword state'];
  var kwUpload = AdWordsApp.bulkUploads().newCsvUpload(header,{'moneyInMicros': false});
   
  var header = ['Campaign', 'Ad group', 'Headline 1', 'Headline 2', 'Description', 'Path 1', 'Path 2', 'Final URL', 'Ad state'];
  var adUpload = AdWordsApp.bulkUploads().newCsvUpload(header,{'moneyInMicros': false});
  
  var total = inputs.length;
  for(var z = 7; z < total; z++) {
    if(!inputs[z][0] || !inputs[z][11] || !inputs[z][10]) { continue; }
    
    var agName = inputs[z][1] + ' - ' + toTitleCase(inputs[z][0]);
    var kw = '[' + inputs[z][0] + ']';
    
    flag = true;
    agUpload.append({
      'Campaign': campaignName, 
      'Ad group': agName, 
      'Ad group state': 'active',
      'Default max. CPC': inputs[z][6]
    });
    
    kwUpload.append({
      'Campaign': campaignName, 
      'Ad group': agName, 
      'Keyword': kw, 
      'Keyword state': 'active'
    });
    
    adUpload.append({
      'Campaign': campaignName, 
      'Ad group': agName, 
      'Keyword': kw, 
      'Headline 1': inputs[z][10], 
      'Headline 2': inputs[1][4], 
      'Description': inputs[z][11], 
      'Path 1': inputs[3][4], 
      'Path 2': inputs[4][4], 
      'Final URL': inputs[z][5], 
      'Ad state': 'active'
    });
  }
  
  if(flag) {
    agUpload.apply();
    Utilities.sleep(5000);
    
    kwUpload.apply();
    adUpload.apply();
  }
  
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    sheet.getRange(7, 1, sheet.getLastRow()-5, sheet.getLastColumn()).clearContent();
    sheet.getRange(1,2).setValue('Export Data');
  }
}

function uploadParams(map, inputs) {
  var campaignName = inputs[4][1];
  if(!campaignName || !AdWordsApp.campaigns().withCondition('Name = "' + campaignName + '"').get().next()) {
    return; 
  }
  
  var iter = AdWordsApp.keywords()
  .withCondition('Status = ENABLED')
  .withCondition('CampaignName = "' + campaignName + '"')
  .get();
  
  while(iter.hasNext()) {
    var kw = iter.next();
    var item = kw.getAdGroup().getName().split(' - ')[0];
    if(!map[item]) { continue; }
    kw.setAdParam(1, map[item].price);
  }
}

function exportConvertingShoppingQueries(map, sheet, inputs) {
  var ids = [];
  var iter = AdWordsApp.shoppingCampaigns().get();
  while(iter.hasNext()) {
    ids.push(iter.next().getId()) 
  }
  
  Logger.log('Reading Keywords in account');  
  
  var existingKeyword = {}, cols = ['Criteria'];
  var reportName = 'KEYWORDS_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where AdGroupStatus = ENABLED and Status = ENABLED',
               'and CampaignStatus IN [ENABLED,PAUSED]',
               'during','YESTERDAY'].join(' ');
  var reportRows = AdWordsApp.report(query, {'includeZeroImpressions': true }).rows();
  while(reportRows.hasNext()) {
    var row = reportRows.next();
    existingKeyword[row.Criteria.replace(/[+]/g, '').toLowerCase()] = 1;
  }
  
  var N = inputs[1][1] ? inputs[1][1] : 30;
  var DATE_RANGE = getAdWordsFormattedDate(N,'yyyyMMdd') + ',' + getAdWordsFormattedDate(1, 'yyyyMMdd');
  
  Logger.log('Reading Search Queries in account');    
  var cols = ['Query','KeywordTextMatchingQuery','Conversions','Cost','ConversionValue'];
  var reportName = 'SEARCH_QUERY_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where CampaignId IN [' + ids.join(',') + ']',
               'and Conversions > 0',
               'and KeywordTextMatchingQuery CONTAINS_IGNORE_CASE "id=="',
               'during',DATE_RANGE].join(' ');
  
  var rows = [];
  var reportRows = AdWordsApp.report(query).rows();
  while(reportRows.hasNext()) {
    var row = reportRows.next();
    if(existingKeyword[row.Query.toLowerCase()]) { continue; }
    
    
    var itemId = row.KeywordTextMatchingQuery.split('id==')[1];
    if(!map[itemId]) {
      continue; 
    }
    
    var headline1 = inputs[0][4].replace('[Brand]', map[itemId].brand) ;
    if(headline1.length > 30) {
      headline1 = inputs[0][8];
    }
    
    var desc = inputs[2][4].replace('[Query]', toTitleCase(row.Query)).replace('[Price]', map[itemId].price) ;
    if(headline1.length > 85) {
      headline1 = inputs[2][8];
    }
    
    rows.push([row.Query, itemId, map[itemId].price, map[itemId].brand, map[itemId].title, map[itemId].url, 
               inputs[2][1], row.Conversions, row.Cost, row.ConversionValue, headline1, desc]);
  }
        
  sheet.getRange(8, 1, sheet.getLastRow()-6, sheet.getLastColumn()).clearContent();
  if(rows.length) {  
    sheet.getRange(8, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function getApprovedShoppingProducts(MERCHANT_ID) {
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
      
      for(var i in products.resources) {
        var product = products.resources[i] 
        var id = product['offerId'];
        map[id] = {
          'price': product['price']['value'], 'brand': product['brand'], 'url': product['link'], 'title': product['title']
        }
      } 
      pageToken = products.nextPageToken;
      pageNum++;
    } while (pageToken);
  } catch(ex){
    Logger.log(MERCHANT_ID + ': ' + ex);
    throw ex;
  }
  
  return map; 
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function toTitleCase(str){
  return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}