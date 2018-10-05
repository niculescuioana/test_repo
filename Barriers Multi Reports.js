var URL = 'https://docs.google.com/spreadsheets/d/1boKOvCn8O3mfa11MLGZ2WDjfKQiaLnMW_-G1WZ9WgAM/edit#gid=60576055';

function main() {
  MccApp.select(MccApp.accounts().withCondition("Name = 'Barriers Direct'").get().next());
  //updateParams();
  
  //return;
  var map = fetchMerchantCentreData();
  
  var hour = parseInt(getAdWordsFormattedDate_(0, 'HH'), 10);

  var day = getAdWordsFormattedDate_(1, 'EEE');
  if(day != 'Sun' && day != 'Sat') {
    compileProductAnomalyReport(map);
  }
  
  compilePriceParams(map);
  compileProductTitleReport(map);
  compileRevenueReport(map);
}

function compileProductAnomalyReport(map) {
  var inputMap = {};
  for(var item in map) {
    inputMap[item] = { 'title': map[item].title, 'price': map[item].price, 'clicks_yesterday': 0, 'clicks_week': 0 };
  }
  
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName', 'AdGroupName', 'Id', 'AdGroupId', 'ProductGroup', 'Clicks'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               'and AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               'during',getAdWordsFormattedDate_(1, 'yyyyMMdd')+','+getAdWordsFormattedDate_(1, 'yyyyMMdd')].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item) { continue; }
    
    if(!inputMap[item]) {
      inputMap[item] = { 'title': '', 'price': '', 'clicks_yesterday': 0, 'clicks_week': 0 };
    }
    
    inputMap[item]['clicks_yesterday'] += parseFloat(row.Clicks);
  }
  
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               'and AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               'during',getAdWordsFormattedDate_(8, 'yyyyMMdd')+','+getAdWordsFormattedDate_(2, 'yyyyMMdd')].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item) { continue; }
    
    if(!inputMap[item]) {
      inputMap[item] = { 'title': '', 'price': '', 'clicks_yesterday': 0, 'clicks_week': 0 };
    }
    
    inputMap[item]['clicks_week'] += parseFloat(row.Clicks);
  }
  
  var out = [['Item Id', 'Title', 'Clicks Yesterday', 'Avg Daily Clicks (Week)', 'Deviation']];
  for(var item in inputMap) {
    var row = inputMap[item];
    row.clicks_week = Math.round(row.clicks_week / 5);
    if(!row.clicks_yesterday && !row.clicks_week) { continue; }
    
    var deviation = row.clicks_week == 0 ? 100 : round_(100*((row.clicks_yesterday - row.clicks_week) / row.clicks_week), 2);
    if(row.clicks_yesterday >= 1.5*row.clicks_week || row.clicks_yesterday <= 0.5*row.clicks_week) {
      out.push([item, row.title, row.clicks_yesterday, row.clicks_week, deviation + '%']); 
    }
  }
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Product Anomalies');
  tab.clearContents();
  
  tab.getRange(1,1,out.length,out[0].length).setValues(out);
  tab.sort(3, false);
  
  if(out.length < 2) { return; }
  
  var EMAIL = 'neeraj@pushgroup.co.uk,mike@pushgroup.co.uk';
  var sub = "Barriers Direct - Product Performance Anomalies";
  var msg = "";  
  
  var htmlBody = '<html><head></head><body>Hi,<br><br>Below is the list of Products which have recieved 50% +/- traffic yesterday compared to weekly average:<br>';   
  htmlBody += 'https://docs.google.com/spreadsheets/d/1boKOvCn8O3mfa11MLGZ2WDjfKQiaLnMW_-G1WZ9WgAM/edit#gid=1968807434';
  htmlBody += '<br><br>Thanks</body></html>';
  
  MailApp.sendEmail(EMAIL, sub, msg, { 'htmlBody': htmlBody});
}

function compilePriceParams(map) {
  //var priceParams = {};
  var out = [];
  for(var id in map) {
    map[id].link = map[id].link.split('?')[0];
    var category = extractCategory(map[id].link);
    
    //if(category && (!priceParams[category] || priceParams[category] > map[id].price)) {
      //priceParams[category] = map[id].price; 
    //}
    out.push([id, map[id].title, map[id].link, category, map[id].price]);
  }
  
  //Logger.log(priceParams['c1117']);
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Price Params');
  tab.getRange(2,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  tab.getRange(2,1,out.length,out[0].length).setValues(out).sort([{'column': 4, 'ascending': true}, {'column': 5, 'ascending': true}]);
  
  updateParams();
}

function updateParams() {
  var priceParams = {};
  var data = SpreadsheetApp.openByUrl(URL).getSheetByName('Price Params').getDataRange().getValues();
  data.shift();
  
  for(var z in data) {
    var category = data[z][3],
        price = data[z][4];
    if(category && (!priceParams[category] || priceParams[category] > price)) {
      priceParams[category] = price; 
    }
  }
  
  var param = {}, noparam = {};
  var query = [
    'SELECT CampaignName, AdGroupName, AdGroupId, Id, CreativeFinalUrls FROM AD_PERFORMANCE_REPORT',
    'WHERE AdGroupStatus = ENABLED and CampaignStatus = ENABLED and Status = ENABLED',
    'and HeadlinePart2 CONTAINS_IGNORE_CASE "param"',
    'and CreativeFinalUrls CONTAINS "h" DURING TODAY'
  ].join(' ');
  var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    var category = extractCategory(JSON.parse(row.CreativeFinalUrls)[0]);
    if(category && priceParams[category]) {
      param[row.AdGroupId] = priceParams[category];
    } else {
      noparam[row.AdGroupId] = [row.CampaignName, row.AdGroupName];
    }
  }
 
  var out = [['Campaign', 'AdGroup']];
  for(var id in noparam) {
    if(!param[id]) {
      out.push(noparam[id]);
    } else {
      delete noparam[id];
    }
  }
  
  if(!AdWordsApp.labels().withCondition('Name = "Missing Price Params"').get().hasNext()) {
    AdWordsApp.createLabel("Missing Price Params");
  }
  
  var ids = Object.keys(param);
  if(ids.length) { 
    var iter = AdWordsApp.ads()
    .withCondition('AdGroupId IN [' + ids.join(',') + ']')
    .withCondition('Status = PAUSED')
    .withCondition('LabelNames CONTAINS_ANY ["Missing Price Params"]')
    .get();
    
    while(iter.hasNext()) {
      var ad = iter.next();
      ad.enable();
      ad.removeLabel('Missing Price Params');
    } 
  }
  
  var iter = AdWordsApp.keywords()
  .withCondition('Status = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .get();
  
  while(iter.hasNext()) {
    var kw = iter.next();
    var price = param[kw.getAdGroup().getId()];
    //Logger.log(kw.getAdGroup().getId + '   :   ' +price);
    if(!price) { continue; }
    
    kw.setAdParam(1, price);
  }
  
  var ids = Object.keys(noparam);
  if(!ids.length) { return; }
  
  var iter = AdWordsApp.ads()
  .withCondition('AdGroupId IN [' + ids.join(',') + ']')
  .withCondition('HeadlinePart2 CONTAINS_IGNORE_CASE "param1"')
  .withCondition('Status = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .get();
  
  while(iter.hasNext()) {
    var ad = iter.next();
    ad.pause();
    ad.applyLabel('Missing Price Params');
  }
  
  var EMAIL = 'neeraj@pushgroup.co.uk,mike@pushgroup.co.uk';
  var sub = "Barriers Direct - Params Not Found";
  var msg = "";  
  
  var htmlBody = '<html><head></head><body>Hi,<br><br>Below is the list of AdGroups which have Param Ads but no prices were found in the feed for them, hence the param ads have been paused and labeled "Missing Price Params":<br><br>';   
  htmlBody += buildTable(out);
  htmlBody += '<br><br>Thanks</body></html>';
  
  MailApp.sendEmail(EMAIL, sub, msg, { 'htmlBody': htmlBody});
}

function extractCategory(link) {
  /*var re = /\d{4,5}/g;
  var matches = link.match(re);
  if(matches && matches[1]) {
   return 'c' + matches[1]; 
  }*/
  
  var parts = link.split('//')[1].split('/');
  if(parts[2] && parts[2].indexOf('-c') > -1) {
    var miniParts = parts[2].split('-c');
    return 'c' + miniParts[miniParts.length-1];
  } else if(parts[1] && parts[1].indexOf('-c') > -1) {
    var miniParts = parts[1].split('-c');
    return 'c' + miniParts[miniParts.length-1];
  }
  
  return '';
}

function compileProductTitleReport(map) {
  var out = [];
  for(var id in map) {
    var title = map[id].title;
    var len = title.length;
    out.push([id, title, len, len > 25 ? title.substring(0,24) : title, len > 70 ? title.substring(0,69) : title ]);
  }
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Product Title Report');
  tab.getRange(2,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
  tab.getRange(2,1,out.length,out[0].length).setValues(out);
}

function compileRevenueReport(map) {
  var TO = getAdWordsFormattedDate_(0, 'yyyyMMdd'); 
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Revenue By Product');
  var data = tab.getDataRange().getValues();
  data.shift();
  
  var exclusionMap = {};
  for(var z in data) {
    if(data[z][8] != 'Yes') { continue; }
    exclusionMap[data[z][0]] = 1;
  }
  
  var results = {}, toReport = [['Campaign', 'AdGroup', 'Item Id', 'Bid']];
  var OPTIONS = { includeZeroImpressions : false };
  var cols = ['CampaignName', 'AdGroupName', 'Id', 'AdGroupId', 'ProductGroup', 'Cost', 'CpcBid'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               "where ProductGroup CONTAINS 'item'",
               "and Cost > 0",
               'during','20161201,'+TO].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var item = extractItemId(row.ProductGroup);
    if(!item) { continue; }
    
    if(exclusionMap[item] && row.CpcBid != "Excluded") {
      toReport.push([row.CampaignName, row.AdGroupName, item, row.CpcBid]); 
    }
    
    if(!map[item]) {
      map[item] = { 'label': '', 'title': '', 'price': 0, 'spend': 0 };
    }
    
    map[item]['spend'] += parseFloat(row.Cost.toString().replace(/,/g,''));
  }
  
  var data = tab.getRange(2,1,tab.getLastRow()-1,4).getValues();
  for(var z in data) {
    var row = map[data[z][0]];
    if(!row ) { continue; }
    
    data[z][1] = row.label;
    data[z][2] = row.title;
    data[z][3] = row.spend;
  }
  
  tab.getRange(2,1,data.length,data[0].length).setValues(data);
  
  if(toReport.length < 2) { return; }
  
  var EMAIL = 'neeraj@pushgroup.co.uk,mike@pushgroup.co.uk';
  var sub = "Barriers Direct - Items Exclusion Errors";
  var msg = "";  
  
  var htmlBody = '<html><head></head><body>Hi,<br><br>Below is the list of items which are marked to be excluded in the Report but have bids in the account:<br><br>';   
  htmlBody += buildTable(toReport);
  htmlBody += '<br><br>Thanks</body></html>';
  
  MailApp.sendEmail(EMAIL, sub, msg, { 'htmlBody': htmlBody});
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
  for(var z in header) {
    table.addHeaderColumn(header[z]);
  }
  
  for(var k in reportData) {
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


function getAdWordsFormattedDate_(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round_(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}




function fetchMerchantCentreData() {
  var map = {};
  var MERCHANT_ID = '2792358';
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
          //Logger.log(products.resources[z]);
          //throw 'www'
          var offerId = products.resources[z].offerId.replace(/ /g,'').toLowerCase();
          map[offerId] = { 
            'title': products.resources[z].title, 'price': parseFloat(products.resources[z].price.value), 
            'spend': 0, 'link': products.resources[z].link ,
            'label': products.resources[z].customLabel0 ,
          };
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

function old() {
  /*
  var URL = 'https://docs.google.com/spreadsheets/d/1U2XTZp-OYudJQ5R0CYJLJ-iGuBkggVK6xoYzEIwSjiQ/edit#gid=15446163';
  var data = SpreadsheetApp.openByUrl(URL).getSheetByName('Sheet1').getDataRange().getValues();
  data.shift();
  
  var map = {};
  for(var z in data) {
  if(!map[data[z][0]]) { map[data[z][0]] = { 'c': 0, 'a': 0} }; 
  
  if(data[z][1]) {
  map[data[z][0]].a += data[z][2];
  } else {
  map[data[z][0]].c += data[z][2]; 
  }
  
  
  }
  
  var out = [];
  for(var camp in map) {
  out.push([camp,map[camp].a,  map[camp].c]);
  }
  
  SpreadsheetApp.openByUrl(URL).getSheetByName('Sheet2').getRange(2,1,out.length,out[0].length).setValues(out);
  return; 
  
  */
  
  
  MccApp.select(MccApp.accounts().withCondition("Name = 'Barriers Direct'").get().next());
  var URL = 'https://docs.google.com/spreadsheets/d/128tuhwiqMpM0GwdiwjxHNSY3M-x3NYnVxTQFCfVWP4U/edit';
  var data = SpreadsheetApp.openByUrl(URL).getSheets()[0].getDataRange().getValues();
  
  data.shift();
  data.shift();
  
  var map = {};
  for(var z in data){
    var url = 'www.barriersdirect.co.uk' + data[z][0].split('?')[0];
    map[url] = 1;
  }
  
  //Logger.log(Object.keys(map).length);
  var START = getAdWordsFormattedDate_(90, 'yyyyMMdd'),
      END = getAdWordsFormattedDate_(0, 'yyyyMMdd');
  
  var output = [];
  var cols = ('FinalUrl,Query,AdGroupName,CampaignName,Cost,Clicks,Impressions,AverageCpc,Ctr,Conversions,CostPerConversion,ConversionRate').split(',');
  var query = [
    'SELECT', cols.join(','),
    'FROM SEARCH_QUERY_PERFORMANCE_REPORT WHERE FinalUrl CONTAINS "http" DURING ', START +','+ END
  ].join(' ');
  
  var rows = AdWordsApp.report(query).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    //Logger.log(row['FinalUrl']);
    
    var finalUrl = row['FinalUrl'].split('?')[0].split('://')[1];
    if(!map[finalUrl]) {
      //Logger.log(finalUrl);
      continue;
    }
    
    var out = [];
    for(var z in cols) {
      out.push(row[cols[z]]); 
    }
    output.push(out);
  }
  
  SpreadsheetApp.openByUrl(URL).getSheets()[1].getRange(2,1,output.length,output[0].length).setValues(output);
}
