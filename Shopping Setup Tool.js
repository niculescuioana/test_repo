
var SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/1hT3o6prboPhBTxU2tFpI0xy2cdaItLCN25L9iYYQW0Q/edit';
var SETTINGS_TAB_NAME = 'Settings';
var ADDITIONAL_FILTERS_TAB_NAME = 'Filters';

// Execution begins here
function main() {
  //if(!AdWordsApp.getExecutionInfo().isPreview()) { return; }
  var INPUTS = parseSettings();
  var ids = Object.keys(INPUTS);
  if(ids.length == 0) { Logger.log('no'); return; }
  
  //ids = ['430-547-6709'];
  MccApp.accounts()
  //.withCondition('Name = "FranksGreatOutdoors"')
  .withIds(ids).executeInParallel('runScript', 'compileResults', JSON.stringify(INPUTS));
}

function parseSettings() {
  var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(SETTINGS_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  var header = ['ACCOUNT_ID','ACCOUNT_NAME','MERCHANT_ID','FLAG','EMAIL',
                'PRODUCT_FILTER_RULE', 'CAMPAIGN_LABEL', 'CAMP_SPLIT_BY', 
                'FALLBACK_CAMP', 'ADGROUP_NAME_TEMPLATE','DEFAULT_BID'];
  
  var SETTINGS = {};
  for(var k in data) {
    if(!data[k][0] || data[k][3] != 'Yes') { continue; }
    var SETTING = { 'ROW_NUM': parseInt(k,10) + 2 };
    for(var j in header) {
      SETTING[header[j]] = data[k][j];
    }
    
    if(!SETTINGS[data[k][0]]) {
      SETTINGS[data[k][0]] = [];
    }
    
    SETTINGS[data[k][0]].push(SETTING);
  }
  
  return SETTINGS;
}

function compileResults() {
  // Do something here
  Logger.log('Finished');
}

function runScript(INPUT) {  
  var INPUTS = JSON.parse(INPUT)[AdWordsApp.currentAccount().getCustomerId()];
  if(!INPUTS || !INPUTS.length) { Logger.log('Missing Settings'); return; }
  
  var approvedShoppingProductsMap = {}, resultsMap = {};
  
  for(var l in INPUTS) {
    var SETTINGS = INPUTS[l];
    SETTINGS.NOW = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss.SSS');
    
    try {
      if(!approvedShoppingProductsMap[SETTINGS.MERCHANT_ID]) {
        approvedShoppingProductsMap[SETTINGS.MERCHANT_ID] = getApprovedShoppingProducts(SETTINGS); 
        resultsMap[SETTINGS.MERCHANT_ID] = fetchMerchantCentreData(SETTINGS);
      }
    } catch(e) {
      var err = e.constructor('Error in Script: ' + e.message);
      err.lineNumber = e.lineNumber - err.lineNumber;
      var errMsg = e+ " (Line: " + err.lineNumber +")"
      var status = [errMsg, SETTINGS.NOW];
      SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(SETTINGS_TAB_NAME).getRange(SETTINGS.ROW_NUM, 13, 1, 2).setValues([status]); 
      throw errMsg;
    }
    
    runForSetting(SETTINGS, approvedShoppingProductsMap[SETTINGS.MERCHANT_ID], resultsMap[SETTINGS.MERCHANT_ID]);
    if(shouldExitNow()) { break; }
  }
}

function runForSetting(SETTINGS, approvedShoppingProducts, results) {
  var status;
  try {
    status = execute(SETTINGS, approvedShoppingProducts, results);
  } catch(e) {
    var err = e.constructor('Error in Script: ' + e.message);
    err.lineNumber = e.lineNumber - err.lineNumber;
    var errMsg = e+ " (Line: " + err.lineNumber +")"
    var status = [errMsg, SETTINGS.NOW];
    SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(SETTINGS_TAB_NAME).getRange(SETTINGS.ROW_NUM, 13, 1, 2).setValues([status]);
    throw errMsg;
  }
  
  SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(SETTINGS_TAB_NAME).getRange(SETTINGS.ROW_NUM, 13, 1, 2).setValues([status]); 
}

function execute(SETTINGS, approvedShoppingProducts, results) {
  
  SETTINGS.FAILED_AG = 0;
  SETTINGS.FILTER_CONDITION = '';
  if(SETTINGS.PRODUCT_FILTER_RULE) {
    fetchFilterRules(SETTINGS); 
    info(SETTINGS.FILTER_CONDITION);
  }
  
  filterShoppingCampaigns(SETTINGS);
  
  if(!SETTINGS.CAMP_IDS.length) { 
    info('No Active Campaign'); 
    return (['No Active Campaign', SETTINGS.NOW]);
  }
  
  
  SETTINGS.NEW_DATA = {};
  SETTINGS.ADGROUP_MAP = {};
  
  var fallBackId = '';
  if(SETTINGS.CAMP_SPLIT_BY) {
    fallBackId = getFallBackCampaign(SETTINGS);
    //filterAllShoppingCampaigns(SETTINGS);
  }
  
  fetchExistingAdGroups(SETTINGS, SETTINGS.CAMP_IDS);
  
 // info(Object.keys(SETTINGS.CAMP_MAP));
  
  SETTINGS.AG_COUNT = 0, SETTINGS.PG_COUNT = 0, SETTINGS.COUNT_ROWS = 0;
  
  var splitCampaignMap = {}, toEnable = [];  
  for(var i in results) {
    if(results[i]['availability'].toLowerCase() != 'in stock') { continue; }  
    if(!approvedShoppingProducts[results[i]['id']]) { continue; }
    
    if(SETTINGS.FILTER_CONDITION && eval(SETTINGS.FILTER_CONDITION)) {
      //info('Skipping due to filter: ' + results[i]);
      continue;
    }
    
    /*if(results[i].offerId != 'GRAFPROLAURED65-R-N-335-46') {
      continue;
    }*/
    
    //info(JSON.stringify(results[i]));
    
    SETTINGS.COUNT_ROWS++;
    
    var campName = SETTINGS.CAMP_SPLIT_BY.toString();
    var agName = SETTINGS.ADGROUP_NAME_TEMPLATE.toString();
    var offerId = results[i].offerId;
    
    for(var attr in results[i]) {
      agName = agName.replace('{' + attr + '}', results[i][attr]);
      campName = campName.replace('{' + attr + '}', results[i][attr]);
    }

    offerId = offerId.toLowerCase();
    //if(offerId != '757348') { continue; }
    
    //Logger.log( SETTINGS.CAMP_IDS);
    var ids = [], toEnableIds = [];
    for(var j in SETTINGS.CAMP_IDS) {
      if(!SETTINGS.ADGROUP_MAP[SETTINGS.CAMP_IDS[j]]) {
        ids.push(SETTINGS.CAMP_IDS[j]);
      } else {
        var agRow = SETTINGS.ADGROUP_MAP[SETTINGS.CAMP_IDS[j]][offerId];
        if(!agRow) {
          ids.push(SETTINGS.CAMP_IDS[j]);
        } else if(agRow.AdGroupStatus == 'paused') {
          toEnableIds.push(agRow.AdGroupId);
        }
      }
    }
    
    //Logger.log(ids.length);
    if(ids.length == 0) { 
      if(toEnableIds.length > 0) {
        toEnable = toEnable.concat(toEnableIds); 
      }
      continue; 
    }
    
    var splitIds = [];
    var splitCampaigns = [];
    if(SETTINGS.CAMP_SPLIT_BY) {
      if(campName) {
        for(var name in SETTINGS.CAMP_MAP) {
          if(name.toLowerCase().indexOf(campName.toLowerCase()) > -1) {
            splitCampaigns.push(SETTINGS.CAMP_MAP[name]);
          }
        }
      }
      
      //Logger.log(fallBackId);
      
      if(splitCampaigns.length == 0) {
        if(fallBackId && (!SETTINGS.ADGROUP_MAP[fallBackId] || !SETTINGS.ADGROUP_MAP[fallBackId][offerId])) {
          if(!SETTINGS.NEW_DATA[fallBackId]) { 
            SETTINGS.NEW_DATA[fallBackId] = {};
          }
          
          SETTINGS.NEW_DATA[fallBackId][agName] = offerId;
        } 
        
        continue;
      } else {
        ids = [], toEnableIds = [];
        for(var j in splitCampaigns) {
          if(!SETTINGS.ADGROUP_MAP[splitCampaigns[j]]) {
            ids.push(splitCampaigns[j]);
          } else {
            var agRow = SETTINGS.ADGROUP_MAP[splitCampaigns[j]][offerId];
            if(!agRow) {
              ids.push(splitCampaigns[j]);
            } else if(agRow.AdGroupStatus == 'paused') {
              toEnable.push(agRow.AdGroupId);
            } 
          }
        }
      }
    }
    
    if(toEnableIds.length > 0) {
      toEnable = toEnable.concat(toEnableIds); 
    }
    
    //Logger.log(ids);
    
    for(var k in ids) {
      if(!SETTINGS.NEW_DATA[ids[k]]) { 
        SETTINGS.NEW_DATA[ids[k]] = {};
      }
      
      SETTINGS.NEW_DATA[ids[k]][agName] = offerId;
    }
   
  }
  
  if(toEnable.length) {
    enableAdGroups(toEnable); 
  }
  
  var campIds = Object.keys(SETTINGS.NEW_DATA);
  if(campIds.length == 0) {
    return (['Success. ' + SETTINGS.COUNT_ROWS + ' rows read. No new products.', SETTINGS.NOW]);
  }
  
  var didExitEarly = false;
  var camps = AdWordsApp.shoppingCampaigns().withIds(campIds).get();
  while(camps.hasNext()) {
    var camp = camps.next();
    var id = camp.getId();
    
    for(var agName in SETTINGS.NEW_DATA[id]) {
      var sku = SETTINGS.NEW_DATA[id][agName];
      var adGroup = createAdGroup(camp, agName, SETTINGS);
      if(!adGroup) { continue; }
      
      buildProductGroups(adGroup, sku, SETTINGS);
      
      if(shouldExitNow()) { didExitEarly = true; break; }
    }
    
    if(didExitEarly) { break; }
  }
  
  var msg = 'Success. ' + SETTINGS.COUNT_ROWS + ' rows read. ' + SETTINGS.AG_COUNT + ' AdGroups Created. ' + SETTINGS.PG_COUNT + ' Product Groups Created.';
  if(SETTINGS.FAILED_AG) {
     msg += ' ' + SETTINGS.FAILED_AG + ' AdGroup Creation Operations Failed.';
    
    if(SETTINGS.EMAIL) {
      var emailMsg = 'Hi,\n\n' + SETTINGS.FAILED_AG + ' AdGroup creation operations have failed in your Account. Please check script logs for more details.\n\nThanks';
      MailApp.sendEmail(SETTINGS.EMAIL, AdWordsApp.currentAccount().getName() + ' Failed Operations in Shopping Update tool', emailMsg);
    }
  }
  
  return ([msg, SETTINGS.NOW]);
}

function shouldExitNow() {
  return (AdWordsApp.getExecutionInfo().getRemainingTime() < (60*2));
}

function enableAdGroups(ids) {
  var iter = AdWordsApp.shoppingAdGroups().withIds(ids).get();
  while(iter.hasNext()) {
    iter.next().enable(); 
  }
}

function buildProductGroups(adGroup, sku, SETTINGS) {
  var productGroup = adGroup.rootProductGroup();
  
  var builder = productGroup.newChild().itemIdBuilder().withValue(sku);
  if(SETTINGS.DEFAULT_BID) {
    builder.withBid(SETTINGS.DEFAULT_BID)
  }
  
  var builderOp = builder.build();
  
  if(!builderOp.isSuccessful()) {
    return;
    //SETTINGS.UNSU 
  }
  
  var pgIter = adGroup.productGroups().get()
  while(pgIter.hasNext()) {
    var pg = pgIter.next();
    if(!pg.isOtherCase()) { continue; }
    pg.exclude();
  } 
  
  //productGroup.newChild().itemIdBuilder().withValue('*').build().getResult().exclude();  
  
  SETTINGS.PG_COUNT++; // += 2;
}

function getFallBackCampaign(SETTINGS) {
  if(!SETTINGS.FALLBACK_CAMP) { return ''; }
  
  var iter = AdWordsApp.shoppingCampaigns()
  .withCondition('Name = "' + SETTINGS.FALLBACK_CAMP + '"')
  .withCondition('Status = ENABLED')
  .get();
  
  while(iter.hasNext()) {
    var camp = iter.next();
    return camp.getId();
  }
}

function createAdGroup(camp, agName, SETTINGS) {
  
  // Build ad group.
  var adGroupOp = camp.newAdGroupBuilder()
  .withName(agName)
  //withMaxCpc(SETTINGS.DEFAULT_BID)
  .build();
  
  // Check for errors.
  if (!adGroupOp.isSuccessful()) {
    SETTINGS.FAILED_AG++;
    return null;
  }
  
  SETTINGS.AG_COUNT++;
  var adGroupResult = adGroupOp.getResult();
  var adGroupAdOp = adGroupResult.newAdBuilder().build();
  
  adGroupResult.createRootProductGroup(); 
  
  return adGroupResult;
}

function fetchExistingAdGroups(SETTINGS, CAMP_IDS) {
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['CampaignName','CampaignId','AdGroupName','AdGroupId','AdGroupStatus'];
  var reportName = 'ADGROUP_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',reportName,
               'where AdGroupStatus IN [ENABLED,PAUSED]',
               SETTINGS.CAMP_SPLIT_BY ? '' : 'and CampaignId IN [' + CAMP_IDS.join(',') + ']',
               'during','TODAY'].join(' ');
  
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(!SETTINGS.ADGROUP_MAP[row.CampaignId]) {
      SETTINGS.ADGROUP_MAP[row.CampaignId] = {}; 
    }
    
    var skuParts = row.AdGroupName.split('>');
    var sku = skuParts[skuParts.length-1].toLowerCase().trim();
    SETTINGS.ADGROUP_MAP[row.CampaignId][sku] = row;
  }
}


function filterShoppingCampaigns(SETTINGS) {
  SETTINGS.CAMP_IDS = [];
  SETTINGS.CAMP_MAP = {};
  var iter = AdWordsApp.shoppingCampaigns().withCondition('Status = ENABLED');
  
  if(SETTINGS.CAMPAIGN_LABEL) {
    iter.withCondition('LabelNames CONTAINS_ANY ["' + SETTINGS.CAMPAIGN_LABEL + '"]') 
  }
  
  iter = iter.get();
  while(iter.hasNext()) {
    var camp = iter.next(); 
    SETTINGS.CAMP_IDS.push(camp.getId()); 
    SETTINGS.CAMP_MAP[camp.getName()] = camp.getId();
  }
}

function fetchMerchantCentreData(SETTINGS) {
  /*var data = DriveApp.getFilesByName('MerchantFile.json').next().getBlob().getDataAsString();
  var results = JSON.parse(data);
  
  return results;*/
  
  var results = [];
  
  var pageToken;
  var pageNum = 1;
  var maxResults = 5;
  
  try {
    // List all the products for a given merchant.
    do {
      var products = ShoppingContent.Products.list(SETTINGS.MERCHANT_ID, {
        pageToken: pageToken,
        maxResults: maxResults
      });
      if (products.resources) {
        results = results.concat(products.resources);
      } 
      pageToken = products.nextPageToken;
      pageNum++;
      break;
    } while (pageToken);
  } catch(ex){
    Logger.log(SETTINGS.MERCHANT_ID + ': ' + ex);
    throw ex;
  }
  
  //throw eee;
  return results;
}

function getApprovedShoppingProducts(SETTINGS) {

  /*var data = DriveApp.getFilesByName('ApprovedMerchantFile.json').next().getBlob().getDataAsString();
  var map = JSON.parse(data);
  
  return map;
  */
  
  var map = {};
  
  var pageToken;
  var pageNum = 1;
  var maxResults = 5;
  
  try {
    // List all the products for a given merchant.
    do {
      var products = ShoppingContent.Productstatuses.list(SETTINGS.MERCHANT_ID, {
        pageToken: pageToken,
        maxResults: maxResults
      });
      
      for(var i in products.resources) {
        var id = products.resources[i]['productId'];
        for(var k in products.resources[i]['destinationStatuses']) {
          var status =  products.resources[i]['destinationStatuses'][k];
          if(status["destination"] == "Shopping" && status["intention"] != "excluded" && status["approvalStatus"] == "approved") {
            map[id] = 1;
            Logger.log(products.resources[i]);
            break;
          }
        }
      } 
      pageToken = products.nextPageToken;
      pageNum++;
      break;
    } while (pageToken);
  } catch(ex){
    Logger.log(SETTINGS.MERCHANT_ID + ': ' + ex);
    throw ex;
  }
  
  return map; 
}

function fetchFilterRules(SETTINGS) {
  var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName(ADDITIONAL_FILTERS_TAB_NAME).getDataRange().getValues();
  data.shift();
  data.shift();
  
  for(var k in data) {
    if(data[k][0] != SETTINGS.PRODUCT_FILTER_RULE) {
      continue;
    }
    
    SETTINGS.FILTER_CONDITION = makeCondition(data[k][1], data[k][2], data[k][3]);
    
    if(data[k][4] && data[k][5] && data[k][6]) {
      SETTINGS.FILTER_CONDITION += data[k][4] == 'AND' ? ' || ' : ' && ';
      SETTINGS.FILTER_CONDITION += makeCondition(data[k][5], data[k][6], data[k][7]);
    }
    
    break;
  }
  
}

function makeCondition(attr, op, val) {
  var condition = '';
  
  if(op == 'Is') {
    condition +=  '(!results[i].' + attr +  ' || (results[i].' + attr +  ' && results[i].' + attr + ' != "' + val + '"))'; 
  } else if(op == 'Is not') {
    condition += '(results[i].' + attr +  ' && results[i].' + attr + ' == "' + val + '")'; 
  } else if(op == 'Contains') {
    condition += '(!results[i].' + attr +  '|| (results[i].' + attr +  ' && results[i].' + attr + '.indexOf("' + val + '") < 0))'; 
  } else if(op == 'Does Not Contain') {
    condition += '(results[i].' + attr +  ' && results[i].' + attr + '.indexOf("' + val + '") > -1)'; 
  } else if(op == 'Greater Than') {
    condition += '(results[i].' + attr +  ' && results[i].' + attr + ' <= ' + val + ')'; 
  } else if(op == 'Less Than') {
    condition += '(results[i].' + attr +  ' && results[i].' + attr + ' >= ' + val + ')'; 
  }
  
  return condition;
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function info(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}