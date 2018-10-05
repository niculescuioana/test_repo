var ID = 98079839, MERCHANT_ID = 112872187;
var URL = 'https://docs.google.com/spreadsheets/d/1DDjTa_y0N5_HIutTiyjEbceOltuKabUsiB-BlB9nAQI/edit';
var FULL_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1lCv9JPk3UNdJdaS1s11KsddChlu5Zcr82Eq4lavlNtM/edit';
var EMAIL = 'cas.paton@onbuy.com,karishma.patel@onbuy.com ';
var CC = ['neeraj@pushgroup.co.uk','charlie@pushgroup.co.uk','ricky@pushgroup.co.uk',
          'ian@pushgroup.co.uk','jay@pushgroup.co.uk']

function main() {
  MccApp.accounts().withCondition('Name = "OnBuy"').executeInParallel('run');
}

function run() {
  compileSummary();
  //compileNonConvertingQueryReport();
  //compileItemIdReport();
  
  //sendEmail();
}

function sendEmail() {
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Email');
  
  
  var SUBJECT = 'OnBuy - Performance Report'; 
  
  var data = tab.getRange(1, 1, 7, 2).getValues();
  var htmlBody = '<html><head></head><body>Hi,<br><br>Please see below an update on performance by channel:<br><br>';    
  htmlBody += buildTable(data); 
  
  var data = tab.getRange(9, 1, 10, 3).getValues();
  htmlBody += '<br><br>Below is an update on Shopping performance:<br><br>';    
  htmlBody += buildTable(data); 
  
  
  var data = tab.getRange(22, 1, 11, 5).getValues();
  htmlBody += '<br><br>Below is the report for Top 10 (By spend) Non Converting Shopping Products (Last 30 Days):<br><br>';
  htmlBody += buildTable(data); 
  
  var data = tab.getRange(36, 1, 11, 6).getValues();
  htmlBody += '<br><br>Below is the report for Top 10 (By spend) 9% Products with OnBuy % sale > 9%:<br><br>';
  htmlBody += buildTable(data); 
  
  var data = tab.getRange(50, 1, 11, 6).getValues();
  if(data[1][0]) {
    htmlBody += '<br><br>Below is the report for Top 10 (By spend) 6% Products with OnBuy % sale > 6%:<br><br>';
    htmlBody += buildTable(data); 
  }
  
  htmlBody += '<br><br>Summary report is available at below url:<br>'+URL;
  
  htmlBody += '<br><br>Detailed report is available at below url:<br>'+FULL_REPORT_URL;
  
  htmlBody += '<br><br>Thanks<br>Push Group</body></html>';
  
  //MailApp.sendEmail('naman@pushgroup.co.uk', SUBJECT, '', { 'htmlBody': htmlBody });
  MailApp.sendEmail(EMAIL, SUBJECT, '', { 'htmlBody': htmlBody, 'cc': CC.join(',') });
}

function compileNonConvertingQueryReport() {
  Logger.log('Fetching Non Converting Queries');  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','AdGroupName', 'Query', 'Cost','Impressions','Clicks'];
  var report = 'SEARCH_QUERY_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where Clicks > 1 and Conversions < 1',
               'during','LAST_30_DAYS'].join(' ');
  
  var rows = [];
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    rows.push([row.CampaignName, row.AdGroupName, row.Query, row.Impressions, row.Clicks, row.Cost]);
  }
  
  var ss = SpreadsheetApp.openByUrl(URL);
  var tab = ss.getSheetByName('Non Converting Queries');
  tab.getRange(2,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  if(rows.length) {
    tab.getRange(2,1,rows.length,rows[0].length).setValues(rows).sort([{'column':1, 'ascending':true}, {'column':2, 'ascending':true}, {'column':6, 'ascending':false}]);
  }
}

function compileItemIdReport() {
  Logger.log('Fetching Merchant Center Data');
  var map = fetchMerchantCentreData();
  
  var stats = {
    'shopping': {},
    '6%': {},
    '9%': {}
  };
  
  Logger.log('Fetching Product Stats');  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName','Id', 'ProductGroup', 'Cost', 'ConversionValue', 'Clicks'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where Conversions < 1',
               "and ProductGroup CONTAINS 'item'",
               'and CampaignName DOES_NOT_CONTAIN "%"',
               'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item || !map[item]) { continue; }
    
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    //row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    
    //var percent_sale = row.Cost / row.ConversionValue;
    
    if(row.CampaignName.indexOf('PLA') > -1 || row.CampaignName.indexOf('GS') > -1) {
      stats['shopping'][item] = { 'Cost': row.Cost, 'Clicks': row.Clicks, 'pct_sale': '' };
    } 
  }
  
  query = ['select',cols.join(','),'from',report,
           'where Conversions > 0',
           "and ProductGroup CONTAINS 'item'",
           'and CampaignName CONTAINS "%"',
           'during','LAST_30_DAYS'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item || !map[item]) { continue; }
    
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g,''));
    row.ConversionValue = parseFloat(row.ConversionValue.toString().replace(/,/g,''));
    
    var percent_sale = row.Cost / row.ConversionValue;
    if(row.CampaignName.indexOf('6%') > -1 && percent_sale > 0.06) {
      stats['6%'][item] = { 'Cost': row.Cost, 'Clicks': row.Clicks, 'pct_sale': percent_sale };
    } else if(row.CampaignName.indexOf('9%') > -1 && percent_sale > 0.09) {
      stats['9%'][item] = { 'Cost': row.Cost, 'Clicks': row.Clicks, 'pct_sale': percent_sale };
    }
  }
  
  var ss = SpreadsheetApp.openByUrl(FULL_REPORT_URL);
  
  Logger.log('Compiling Reports');
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Summary');
  var items = stats['shopping'];
  
  tab.getRange(23,1,10,6).clearContent();
  var out = getRows(map, items);
  if(out.length) {
    var sheet = ss.getSheetByName('Shopping - Non Converting');
    sheet.getRange(4,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
    sheet.getRange(4,1,out.length,out[0].length).setValues(out);
    
    if(out.length > 10) { out.length = 10; }
    tab.getRange(23,1,out.length,out[0].length).setValues(out);
  }
  
  items = stats['6%'];
  tab.getRange(51,1,10,6).clearContent();
  if(Object.keys(items).length > 0) {
    out = getRows(map, items);
    if(out.length) {
      var sheet = ss.getSheetByName('6% Products (OnBuy % sale > 6%)');
      sheet.getRange(4,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
      sheet.getRange(4,1,out.length,out[0].length).setValues(out);
      
      if(out.length > 10) { out.length = 10; }
      tab.getRange(51,1,out.length,out[0].length).setValues(out);
    }
  }
  
  items = stats['9%'];
  out = getRows(map, items);
  tab.getRange(37,1,10,6).clearContent();
  if(out.length) {
    
    var sheet = ss.getSheetByName('9% Products (OnBuy % sale > 9%)');
    sheet.getRange(4,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
    sheet.getRange(4,1,out.length,out[0].length).setValues(out);
    
    out.length = 10;
    tab.getRange(37,1,out.length,out[0].length).setValues(out);
  } 
}

function getRows(map, items) {
  var keys = Object.keys(items);
  keys.sort(function(a,b){ return items[b].Cost - items[a].Cost});
  
  var out = [];
  for(var z=0; z < keys.length; z++) {
    out.push([map[keys[z]].title, keys[z], items[keys[z]].Cost, items[keys[z]].Clicks, map[keys[z]].price, items[keys[z]].pct_sale]);
  }
  
  return out;
}

function compileSummary() {
  var end = getAdWordsFormattedDate_(0, 'yyyy-MM-dd');
  var start = end.substring(0,8) + '01';
  
  var stats = { 'searchSpend': 0, 'displaySpend': 0, 'videoSpend': 0, 'shoppingSpend': 0 };
  var query = [
    'SELECT CampaignId, AdvertisingChannelType, Cost FROM CAMPAIGN_PERFORMANCE_REPORT DURING THIS_MONTH'
  ].join(' ');
  
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
   var row = rows.next();
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g, ''));
    
    if(row.AdvertisingChannelType == 'Video') {
      stats['videoSpend'] += row.Cost;
    } else if(row.AdvertisingChannelType == 'Shopping') {
      stats['shoppingSpend'] += row.Cost;
    } else if(row.AdvertisingChannelType == 'Display') {
      stats['displaySpend'] += row.Cost;
    } else {
      stats['searchSpend'] += row.Cost;
    }
  }
  /*
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = SEARCH').get();
  while(iter.hasNext()) {
    stats.searchSpend +=  iter.next().getStatsFor('THIS_MONTH').getCost();
  }
  
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = DISPLAY').get();
  while(iter.hasNext()) {
    stats.displaySpend +=  iter.next().getStatsFor('THIS_MONTH').getCost();
  }
  
  var iter = AdWordsApp.shoppingCampaigns().get();
  while(iter.hasNext()) {
    stats.shoppingSpend +=  iter.next().getStatsFor('THIS_MONTH').getCost();
  }
  
  var iter = AdWordsApp.videoCampaigns().get();
  while(iter.hasNext()) {
    stats.videoSpend +=  iter.next().getStatsFor('THIS_MONTH').getCost();
  }
  */
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Summary');
  tab.getRange(2,2,4,1).setValues([
    [stats.shoppingSpend], [stats.searchSpend],
    [stats.displaySpend], [stats.videoSpend]
  ]);
  
  var optargs = {'samplingLevel': 'HIGHER_PRECISION', 'dimensions':'ga:campaign','max-results': 10000 }
  var stats_9 = {'Transactions': 0, 'Cost': 0, 'Clicks': 0, 'Revenue': 0},
      stats_6 = {'Transactions': 0, 'Cost': 0, 'Clicks': 0, 'Revenue': 0};
  
  getStatsFromAnalytics(start, end, stats_6, stats_9, optargs);
  
  
  tab.getRange(10,2,2,2).setValues([[stats_6.Cost, stats_9.Cost],[stats_6.Clicks, stats_9.Clicks]]);
  
  tab.getRange(15,2,2,2).setValues([[stats_6.Transactions, stats_9.Transactions],[stats_6.Revenue, stats_9.Revenue]]);
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
  Logger.log(rows);
  
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




function buildTable(reportData) {
  var table = new HTMLTable();
  table.setTableStyle(['font-family: "Lucida Sans Unicode","Lucida Grande",Sans-Serif;',
                       'font-size: 12px;',
                       'background: #fff;',
                       'margin: 45px;',
                       'width: 480px;',
                       'border-collapse: collapse;',
                       'text-align: left'].join(''));
  table.setHeaderStyle(['font-size: 14px;',
                        'font-weight: normal;',
                        'color: #039;',
                        'padding: 10px 8px;',
                        'border-bottom: 2px solid #6678b1'].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  
  var header = reportData.shift();
  for(var k in header) {
    table.addHeaderColumn(header[k]);
  }  
  
  for(var k in reportData) {
    if(!reportData[k][0]) { continue; }
    table.newRow();
    for(var j in reportData[k]){
      table.addCell(reportData[k][j]);
    }  
  }
  return table.toString();
}


/*********************************************
* HTMLTable: A class for building HTML Tables
* Version 1.0
**********************************************/
function HTMLTable() {
  this.headers = [];
  this.columnStyle = {};
  this.body = [];
  this.currentRow = 0;
  this.tableStyle;
  this.headerStyle;
  this.cellStyle;
  
  this.addHeaderColumn = function(text) {
    this.headers.push(text);
  };
  
  this.addCell = function(text,style) {
    if(!this.body[this.currentRow]) {
      this.body[this.currentRow] = [];
    }
    this.body[this.currentRow].push({ val:text, style:(style) ? style : '' });
  };
  
  this.newRow = function() {
    if(this.body != []) {
      this.currentRow++;
    }
  };
  
  this.getRowCount = function() {
    return this.currentRow;
  };
  
  this.setTableStyle = function(css) {
    this.tableStyle = css;
  };
  
  this.setHeaderStyle = function(css) {
    this.headerStyle = css; 
  };
  
  this.setCellStyle = function(css) {
    this.cellStyle = css;
    if(css[css.length-1] !== ';') {
      this.cellStyle += ';';
    }
  };
  
  this.toString = function() {
    var retVal = '<table ';
    if(this.tableStyle) {
      retVal += 'style="'+this.tableStyle+'"';
    }
    retVal += '>'+_getTableHead(this)+_getTableBody(this)+'</table>';
    return retVal;
  };
  
  function _getTableHead(instance) {
    var headerRow = '';
    for(var i in instance.headers) {
      headerRow += _th(instance,instance.headers[i]);
    }
    return '<thead><tr>'+headerRow+'</tr></thead>';
  };
  
  function _getTableBody(instance) {
    var retVal = '<tbody>';
    for(var r in instance.body) {
      var rowHtml = '<tr>';
      for(var c in instance.body[r]) {
        rowHtml += _td(instance,instance.body[r][c]);
      }
      rowHtml += '</tr>';
      retVal += rowHtml;
    }
    retVal += '</tbody>';
    return retVal;
  };
  
  function _th(instance,val) {
    var retVal = '<th scope="col" ';
    if(instance.headerStyle) {
      retVal += 'style="'+instance.headerStyle+'"';
    }
    retVal += '>'+val+'</th>';
    return retVal;
  };
  
  function _td(instance,cell) {
    var retVal = '<td ';
    if(instance.cellStyle || cell.style) {
      retVal += 'style="';
      if(instance.cellStyle) {
        retVal += instance.cellStyle;
      }
      if(cell.style) {
        retVal += cell.style;
      }
      retVal += '"';
    }
    retVal += '>'+cell.val+'</td>';
    return retVal;
  };
}