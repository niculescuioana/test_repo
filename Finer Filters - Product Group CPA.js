function main(){
  MccApp.accounts().withCondition('Name = "Finer Filters"').executeInParallel('run');
}

function run() {
  var PROFILE_ID = 92590400;
  
  var url = 'https://docs.google.com/spreadsheets/d/1Xy4UanKOyHQodVWmpYJzIP_NiHTitUQ_sqeDV1JYOoI/edit';
  var inputSheet = SpreadsheetApp.openByUrl(url).getSheetByName('Product Group Report');
  
  var SCRIPT_NAME = 'PRODUCT GROUP REPORT';  
  var scriptNameHeader = inputSheet.getRange(1,1,1,inputSheet.getLastColumn()).getValues()[0];
  var columnIndex = scriptNameHeader.indexOf(SCRIPT_NAME)+1;
  
  var headers = ['NAME', 'FLAG', 'REPORT_TYPE', 'CPA_REPORT_URL', 'ROAS_REPORT_URL'];
  
  var SETTINGS = new Object();	 
  
  var rowNum = -1;
  var accName = AdWordsApp.currentAccount().getName();
  var data = inputSheet.getDataRange().getValues();
  for(var z in data) {
    if(data[z][0] == accName) {
      for(var j in data[z]){
        SETTINGS[headers[j]] = data[z][j];
      }
      
      rowNum = parseInt(z, 10) + 1;
      break; 
    }
  }
  
  //Logger.log(SETTINGS);
  
  //if(SETTINGS.FLAG.toUpperCase() != 'Y'){ info('Script turned off for the account. Exiting.'); return; }
  
  SETTINGS.PROFILE_ID = PROFILE_ID;
  var urlIndex = columnIndex + headers.length - 2;
  compileCPAReport(SETTINGS, urlIndex);
  
  var dateString = getAdWordsFormattedDate(0, 'MMM dd, yyyy HH:mm');
  var index = columnIndex + headers.length - 1;
  inputSheet.getRange(rowNum,index,1,1).setValue(dateString);
}

function compileCPAReport(SETTINGS, urlIndex) {
  var accName = AdWordsApp.currentAccount().getName();
  var reportName = "Product Group CPA Report - " + accName;
  
  var ss, createSpreadsheet = false;
  try {
    ss = SpreadsheetApp.openByUrl(SETTINGS.CPA_REPORT_URL);
  } catch(exep) {
    createSpreadsheet = true;
  }   
  
  if(!ss || createSpreadsheet) {
    return
  } 
  
  var PROFIT_MAP = parseInputSheet(SETTINGS.CPA_REPORT_URL);
  
  
  var TAB_NAME = 'MTD';
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  var FROM = TO.substring(0,8) + '01';
  var MTD =  FROM.replace(/-/g, '') + ',' + TO.replace(/-/g, '');
  compileReport('CPA', TAB_NAME, FROM, TO, MTD, PROFIT_MAP, SETTINGS);
  
  
  var TAB_NAME = 'Last 7 Days';
  var FROM = getAdWordsFormattedDate(7, 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var LAST_7_DAYS =  FROM.replace(/-/g, '') + ',' + TO.replace(/-/g, '');
  compileReport('CPA', TAB_NAME, FROM, TO,  LAST_7_DAYS, PROFIT_MAP, SETTINGS);
  
  
  var TAB_NAME = 'Last 90 Days';
  var FROM = getAdWordsFormattedDate(90, 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var LAST_90_DAYS =  FROM.replace(/-/g, '') + ',' + TO.replace(/-/g, '');
  compileReport('CPA', TAB_NAME, FROM, TO, LAST_90_DAYS, PROFIT_MAP, SETTINGS);
  
  var TAB_NAME = 'Last 30 Days';
  var FROM = getAdWordsFormattedDate(30, 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(1, 'yyyy-MM-dd');
  var LAST_30_DAYS =  FROM.replace(/-/g, '') + ',' + TO.replace(/-/g, '');
  var pgMap = compileReport('CPA', TAB_NAME, FROM, TO,  LAST_30_DAYS, PROFIT_MAP, SETTINGS);
  
  manageBids(pgMap, PROFIT_MAP);
}  


function parseInputSheet(URL) {
  var SETTINGS = {};
  var rows = SpreadsheetApp.openByUrl(URL).getSheetByName('Inputs').getDataRange().getValues();
  rows.shift();
  
  for(var k in rows) {
    if(!rows[k][0]) { continue; }
    SETTINGS[rows[k][0]] =  { 
      'NAME': rows[k][1], 'PROFIT': rows[k][2],
      'MIN_CLICKS': rows[k][3], 'CPA_ABOVE': rows[k][4], 'DEC_PCT': rows[k][5], 
      'MIN_BID': rows[k][6], 'CPA_BELOW': rows[k][7], 'INC_PCT': rows[k][8], 'MAX_BID': rows[k][9]
    }
  }
  
  return SETTINGS;
}

function compileReport(key, TAB_NAME, FROM, TO, DATE_RANGE, PROFIT_MAP, SETTINGS) {
  var OPTIONS = { includeZeroImpressions : false };
  
  var pgMap = {};
  var cols = ['AdGroupId','Id','ProductGroup','CpcBid','Clicks','Impressions',
              'Cost', 'Conversions','CostPerConversion','ConversionValue'];
  var reportName = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where CampaignStatus = ENABLED and AdGroupStatus = ENABLED',
               'during',DATE_RANGE].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var itemId = extractItemId(row.ProductGroup);
    if(!itemId) { continue; }
    
    var map = PROFIT_MAP[itemId];
    if(!map) {
      map = { 'NAME': '', 'PROFIT': 0 } 
    }
    
    if(!pgMap[itemId]) {
      pgMap[itemId] = { 
        'AdGroupId': row.AdGroupId, 'Id': row.Id, 'CpcBid': row.CpcBid,
        'ProductGroup': row.ProductGroup, 'ProductGroupName': map.NAME, 'Profit' : map.PROFIT, 
        'Clicks': 0,  'Impressions': 0, 'Cost': 0,  'Conversions': 0, 'ConversionValue': 0,
        'AssistedConversions': 0, 'AssistedConversionValue': 0
      }
    }
    
    pgMap[itemId].Clicks += parseInt(row.Clicks,10);
    pgMap[itemId].Impressions += parseInt(row.Impressions,10);
    pgMap[itemId].Conversions += parseInt(row.Conversions,10);
    pgMap[itemId].Cost += parseFloat(row.Cost.toString().replace(/,/g,''));
    pgMap[itemId].ConversionValue += parseFloat(row.ConversionValue.toString().replace(/,/g,''));
  }
  
  
  addAssistedDataFromAnalytics(SETTINGS.PROFILE_ID, pgMap, FROM, TO);
  
  
  var header = [
    'Product Group','Product Name', 
    'Clicks','Impressions','Cost', 'Conversions','CPA',
    'Assisted Conversions', 'Total Conversions', 'Total CPA',
    'Profit', 'Net Profit Per Sale', 'Net Profit'
  ];
  
  
  var headerBg = [];
  var greenBg = [];
  var redBg = [];
  var whiteBg = [];
  
  while(headerBg.length < header.length) {
    headerBg.push('#efefef');
    greenBg.push('#d9ead3');
    redBg.push('#ea9999');
    whiteBg.push('#fff');
  }
  
  var backgrounds = [headerBg];
  var output = [header];
  
  for(var itemId in pgMap) {
    var row = pgMap[itemId];
    row.TotalConversions = row.AssistedConversions + row.Conversions;
    row.TotalCPA = row.TotalConversions == 0 ? row.Cost : round((row.Cost/row.TotalConversions),2);
    row.CostPerConversion = row.Conversions == 0 ? row.Cost : round((row.Cost/row.Conversions),2);
    
    row.TotalRevenue = row.AssistedConversionValue + row.ConversionValue;
    row.TotalROAS = row.Cost == 0 ? 0 : round(parseFloat(row.TotalRevenue / row.Cost),2);
    row.ROAS = row.Cost == 0 ? 0 : round(parseFloat(row.ConversionValue / row.Cost),2);
    
    var netProfit = row.Profit == 0 ? 'N/A' : (row.Conversions * row.Profit) - row.Cost;
    var netProfitPerSale = row.Cost == 0 || row.Conversions == 0 || row.Profit == 0 || isNaN(row.Profit) ? 'N/A' : row.Profit - row.TotalCPA;
    var reportRow = [row.ProductGroup, row.ProductGroupName, row.Clicks, row.Impressions,
                     row.Cost, row.Conversions, row.CostPerConversion, 
                     row.AssistedConversions, row.TotalConversions, row.TotalCPA,
                     row.Profit ? row.Profit : '', netProfitPerSale, netProfit];
    
    
    if(netProfit == 'N/A') {
      backgrounds.push(whiteBg);
    } else if(netProfit > 0) {
      backgrounds.push(greenBg);
    } else if(netProfit < 0) {
      backgrounds.push(redBg);  
    } else {
      backgrounds.push(whiteBg);
    }
    
    output.push(reportRow);
  }
  
  var REP_URL = SETTINGS.CPA_REPORT_URL;
  var sheet = SpreadsheetApp.openByUrl(REP_URL).getSheetByName(TAB_NAME);
  sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).clear();
  
  sheet.getRange(1,1,1,output[0].length).setFontWeight('bold');
  sheet.getRange(1,1,output.length,output[0].length).setValues(output).setBackgrounds(backgrounds);
  sheet.setFrozenRows(1);
  
  sheet.getDataRange().setFontFamily('Calibri').setVerticalAlignment('middle');
  sheet.getRange('C:M').setHorizontalAlignment('center');
  
  return pgMap;
}

function addAssistedDataFromAnalytics(PROFILE_ID, pgMap, FROM, TO) {
  var filters = ['mcf:basicChannelGroupingPath=@Paid Search',
                 'mcf:basicChannelGroupingPath!=Paid Search',
                 'mcf:conversionType==Transaction'];
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:keywordPath',
    'filters': filters.join(';')
  };
  
  var results = Analytics.Data.Mcf.get(
    'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
    FROM,                 // Start-date (format yyyy-MM-dd).
    TO,                  // End-date (format yyyy-MM-dd).
    "mcf:totalConversions,mcf:totalConversionValue",
    optArgs
  );
  
  var rows = results.rows;
  for(var k in rows) {
    var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
    if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
    var index = channelGroups.length-1;
    if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
    
    var id = '';
    var paths = rows[k][1]['conversionPathValue'];
    for(var j in paths) {
      if(paths[j].nodeValue.indexOf('id==') > -1) {
        id = paths[j].nodeValue.replace('id==', '');
        break;
      }
    }
    
    if(pgMap[id]) {
      pgMap[id].AssistedConversions += parseInt(rows[k][2].primitiveValue,10);
      pgMap[id].AssistedConversionValue += parseFloat(rows[k][3].primitiveValue);    
    }
  }
}


function manageBids(pgMap, map) {
  
  var ids = [], toUpdate = {};
  
  for(var id in pgMap) {
    if(!map[id]) { 
      continue; 
    }
    
    var row = pgMap[id];
    var SETTINGS = map[id];
    if(row.Clicks < SETTINGS.MIN_CLICKS) { continue; }
    
    var DEC_FACTOR = 1 - SETTINGS.DEC_PCT;
    var INC_FACTOR = 1 + SETTINGS.INC_PCT;
    row.CpcBid = parseFloat(row.CpcBid);
    //Logger.log(row.TotalCPA + ' : ' + SETTINGS.CPA_ABOVE);
    var newBid = '';
    if(SETTINGS.CPA_ABOVE && row.TotalCPA > SETTINGS.CPA_ABOVE) {
      newBid = row.CpcBid*DEC_FACTOR;
      if(row.CpcBid - newBid  < 0.01) {
        newBid = row.CpcBid - 0.01;
      }
      if(newBid < SETTINGS.MIN_BID) {
        newBid = SETTINGS.MIN_BID;
      }
    } else if(SETTINGS.CPA_BELOW && row.TotalCPA < SETTINGS.CPA_BELOW) {
      newBid = row.CpcBid*INC_FACTOR;
      if(newBid - row.CpcBid < 0.01) {
        newBid = row.CpcBid + 0.01;
      }
      if(newBid > SETTINGS.MAX_BID) {
        newBid = SETTINGS.MAX_BID;
      }
    }
    
    if(!newBid) { continue; }
    var key = [row.AdGroupId, row.Id].join('-');
    ids.push([row.AdGroupId, row.Id]);
    toUpdate[key] = newBid;
  }
  
  var iter =  AdWordsApp.productGroups().withIds(ids).get();
  info(iter.totalNumEntities());
  while(iter.hasNext()) {
    var pg = iter.next(); 
    var bid = round(toUpdate[[pg.getAdGroup().getId(), pg.getId()].join('-')],2);
    try {
      pg.setMaxCpc(bid);
    } catch(ex) {
      Logger.log(toUpdate[[pg.getAdGroup().getId(), pg.getId()].join('-')]);
      throw ex;
    }
  }
}

function extractItemId(str) {
  //return str.match(/(?:"[^"]*"|^[^"]*$)/)[0].replace(/"/g, "");
  var matches = str.match(/item id = \"(.*?)\"/);
  if(matches) { 
    return matches[1];   
  }
  
  return ''
}

function info(msg) {
  var time = Utilities.formatDate(new Date(),AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS');
  Logger.log(time + ' - ' + msg);
}


function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}