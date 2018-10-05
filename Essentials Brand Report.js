var ID = 14426663, MERCHANT_ID = 101722908;
var URL = 'https://docs.google.com/spreadsheets/d/1soSEe5SW1sPcVJUUVJ1QnNzYFo56dQwTIjuY8iNmcug/edit';
var TAB_NAME = '2018 Brands';

function main() {
  MccApp.accounts().withCondition('Name = "Essentials London"').executeInParallel('run');
}


function run() {
  var initMap = {}, agMap = {};
  compileBrandSummary(initMap, agMap);
  compileItemReport(initMap, agMap)
}


function compileItemReport(initMap, agMap) {
  Logger.log('Fetching Merchant Center Data');
  var map = fetchMerchantCentreData();
  
  var results = {};
  var agIds = Object.keys(agMap);
  Logger.log('Fetching Product Stats');  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','Id', 'AdGroupId', 'ProductGroup', 'Cost', 'ConversionValue', 'Clicks', 'Conversions', 'Impressions'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               "and Cost > 0 and ConversionValue > 0",
               "and AdGroupId IN [" + agIds.join(',') + "]",
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item || !map[item]) { continue; }
    
    var brand = agMap[row.AdGroupId];
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    
    var percent_sale = row.ConversionValue > 0 ? row.Cost / row.ConversionValue : 0;
    
    
    if(!results[brand]) {
      results[brand] = []; 
    }
    
    results[brand].push([map[item]['title'], percent_sale, item]);
  }
   
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
  var rowNum = 44;
  tab.getRange(rowNum + 19, 1, tab.getLastRow()-35, tab.getLastColumn()).clear();
  
  for(var brand in initMap) {
    if(!results[brand]) { continue; }
    
    if(rowNum != 44) {
      tab.getRange(44, 1, 18, tab.getLastColumn()).copyFormatToRange(tab, 1, tab.getLastColumn(), rowNum, rowNum+17);
      tab.getRange(45, 1, 2, tab.getLastColumn()).copyTo(tab.getRange(rowNum+1, 1, 2, tab.getLastColumn())); 
    }
    
    tab.getRange(rowNum,1).setValue(brand);
    
    var below = [], above = [];
    var rows = results[brand];
    for(var z in rows) {
      if(rows[z][1] <= 0.07) {
        below.push(rows[z])
      } else {
        above.push(rows[z])
      }
    }
    
    if(below.length > 15) {
      below.length = 15;
    }
    
    if(above.length > 15) {
      above.length = 15;
    }
    
    tab.getRange(rowNum+3,1,15,4).clearContent();
    if(below.length) {
      tab.getRange(rowNum+3, 1, below.length, below[0].length).setValues(below);
    }
    
    tab.getRange(rowNum+3,6,15,4).clearContent();
    if(above.length) {
      tab.getRange(rowNum+3, 6, above.length, above[0].length).setValues(above);
    }
    
    rowNum += 19;
  }
}

function compileBrandSummary(initMap, agMap){
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
  var input = tab.getRange(5,1,15,8).getValues();
  for(var z in input) {
    if(input[z][0]) {
      initMap[input[z][0]] = { 'Cost': 0, 'ConversionValue': 0, 'Conversions': 0, 'Clicks': 0, 'Impressions': 0 };
    }
    
    if(input[z][5]) {
      initMap[input[z][5]] = { 'Cost': 0, 'ConversionValue': 0, 'Conversions': 0, 'Clicks': 0, 'Impressions': 0 };
    }
  }
  
  var brandMap = JSON.parse(JSON.stringify(initMap));
  var query = [
    'SELECT AdGroupId, AdGroupName, Labels, Cost, ConversionValue, Conversions, Clicks, Impressions',
    'FROM ADGROUP_PERFORMANCE_REPORT WHERE Cost > 0 DURING LAST_30_DAYS'
  ].join(' ');
  
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next(); 
    for(var label in brandMap) {
      if(row['Labels'].indexOf(label) > -1) {
        agMap[row['AdGroupId']] = label;
        brandMap[label]['Cost'] += parseFloat(row['Cost'].toString().replace(/,/g,''));
        brandMap[label]['ConversionValue'] += parseFloat(row['ConversionValue'].toString().replace(/,/g,''));
        
        brandMap[label]['Conversions'] += parseFloat(row['Conversions'].toString().replace(/,/g,''));
        brandMap[label]['Clicks'] += parseInt(row['Clicks'], 10);
        brandMap[label]['Impressions'] += parseInt(row['Impressions'], 10);
      }
    }
  }
  
  //Clicks	Impressions	Ctr	Avg Cpc	Conversions	CPA
  var output = [];
  for(var z in input) {
    if(input[z][0]) {
      var row = brandMap[input[z][0]];
      input[z][2] = row['ConversionValue'] == 0 ? 0 : row['Cost'] / row['ConversionValue'];
      
      row['CPC'] = row['Clicks'] == 0 ? 0 : row['Cost'] / row['Clicks'];
      row['CTR'] = row['Impressions'] == 0 ? 0 : row['Clicks'] / row['Impressions'];
      row['CPA'] = row['Conversions'] == 0 ? 0 : row['Cost'] / row['Conversions'];
      
      output.push([
        input[z][0], row['Cost'], row['ConversionValue'], input[z][2], 
        row['Clicks'], row['Impressions'], row['CTR'], row['CPC'], row['Conversions'], row['CPA']
      ]);
      
    }
    
    if(input[z][5]) {
      var row = brandMap[input[z][5]];
      input[z][7] = row['ConversionValue'] == 0 ? 0 : row['Cost'] / row['ConversionValue'];
      
      row['CPC'] = row['Clicks'] == 0 ? 0 : row['Cost'] / row['Clicks'];
      row['CTR'] = row['Impressions'] == 0 ? 0 : row['Clicks'] / row['Impressions'];
      row['CPA'] = row['Conversions'] == 0 ? 0 : row['Cost'] / row['Conversions'];
      
      output.push([
        input[z][5], row['Cost'], row['ConversionValue'], input[z][7], 
        row['Clicks'], row['Impressions'], row['CTR'], row['CPC'], row['Conversions'], row['CPA']
      ]);
    }
  }
  
  tab.getRange(5,1,15,8).setValues(input);
  
  
  var sheet = SpreadsheetApp.openByUrl(URL).getSheetByName('Brand Summary');
  sheet.getRange(3,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
  sheet.getRange(3,1,output.length,output[0].length).setValues(output);
  
  var brandMap = {};
  var query = [
    'SELECT AdGroupId, AdGroupName, Labels, Cost, ConversionValue',
    'FROM ADGROUP_PERFORMANCE_REPORT WHERE Cost > 0 DURING LAST_7_DAYS'
  ].join(' ');
  
  var key = 'L7';
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next(); 
    for(var label in initMap) {
      if(row['Labels'].indexOf(label) > -1) {
        if(!brandMap[label]) {
          brandMap[label] = {
            'L7': { 'Cost': 0, 'ConversionValue': 0 },
            'P7': { 'Cost': 0, 'ConversionValue': 0 }
          }
        }
        
        brandMap[label][key]['Cost'] += parseFloat(row['Cost'].toString().replace(/,/g,''));
        brandMap[label][key]['ConversionValue'] += parseFloat(row['ConversionValue'].toString().replace(/,/g,''));
      }
    }
  }
  
  var START = getAdWordsFormattedDate_(14, 'yyyyMMdd'),
      END = getAdWordsFormattedDate_(8, 'yyyyMMdd');
  
  var query = [
    'SELECT AdGroupId, AdGroupName, Labels, Cost, ConversionValue',
    'FROM ADGROUP_PERFORMANCE_REPORT WHERE Cost > 0 DURING '+START+','+END
  ].join(' ');
  
  var key = 'P7';
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next(); 
    for(var label in initMap) {
      if(row['Labels'].indexOf(label) > -1) {
        if(!brandMap[label]) {
          brandMap[label] = {
            'L7': { 'Cost': 0, 'ConversionValue': 0 },
            'P7': { 'Cost': 0, 'ConversionValue': 0 }
          }
        }
        
        brandMap[label][key]['Cost'] += parseFloat(row['Cost'].toString().replace(/,/g,''));
        brandMap[label][key]['ConversionValue'] += parseFloat(row['ConversionValue'].toString().replace(/,/g,''));
      }
    }
  }
  
  var out = [];
  for(var brand in brandMap) {
    var row = brandMap[brand];
    //if(!row['P7']['ConversionValue']) { continue; }
    
    row['L7'].CPS = row['L7']['ConversionValue'] == 0 ? 0 : row['L7']['Cost'] / row['L7']['ConversionValue'];
    row['P7'].CPS = row['P7']['ConversionValue'] == 0 ? 0 : row['P7']['Cost'] / row['P7']['ConversionValue'];
    
    var diff = (row['L7'].CPS - row['P7'].CPS);
    out.push([brand, row['L7'].CPS, row['P7'].CPS, diff]);
  }
  
  out = out.sort(function(a,b) {return b[3]-a[3];});
  
  var inc = [], dec = [];
  for(var z in out) {
    if(out[z][3] > 0) {
      inc.push(out[z]);
    } else if(out[z][3] < 0) {
      dec.push(out[z]);
    }
  }
  
  if(inc.length > 15) {
    inc.length = 15;
  }
  
  if(dec.length > 15) {
    dec.length = 15;
  }
  
  tab.getRange(25,1,15,4).clearContent();
  
  if(inc.length) {
    tab.getRange(25,1,inc.length,inc[0].length).setValues(inc);
  }
  
  tab.getRange(25,6,15,4).clearContent();
  
  if(dec.length) {
    tab.getRange(25,6,dec.length,dec[0].length).setValues(dec);
  }
  
}

function getStatsFromAnalytics(FROM, TO, stats_6, stats_9, optargs) {
  var results, attempts = 3;
  while(attempts > 0) {
    try {    
      results = Analytics.Data.Ga.get(
        'ga:'+ID,              // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                   // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:adClicks,ga:transactions,ga:transactionRevenue",
        optargs);
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + ID);
      attempts--;
      Utilities.sleep(2000);
    }  
  }
  
  var rows = results.getRows();
  for(var k in rows) {
    if(rows[k][0].indexOf('9%') > -1) {
      stats_9.Transactions += parseInt(rows[k][3],10);
      stats_9.Clicks += parseInt(rows[k][2],10);
      stats_9.Cost += parseFloat(rows[k][1].toString().replace(/,/g,''));
      stats_9.Revenue += parseFloat(rows[k][4].toString().replace(/,/g,''));
    } else if (rows[k][0].indexOf('6%') > -1) {
      stats_6.Transactions += parseInt(rows[k][3],10);
      stats_6.Clicks += parseInt(rows[k][2],10);
      stats_6.Cost += parseFloat(rows[k][1].toString().replace(/,/g,''));
      stats_6.Revenue += parseFloat(rows[k][4].toString().replace(/,/g,''));
    }
  }
}

function getAdWordsFormattedDate_(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,'PST',format);
}

function round_(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}




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

function extractItemId(str) {
  //return str.match(/(?:"[^"]*"|^[^"]*$)/)[0].replace(/"/g, "");
  var matches = str.match(/item id = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}

function extractCustomLabel(str) {
  var matches = str.match(/custom label 1 = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}