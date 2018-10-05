/*************************************************
* Disapproved Call Extensions Alert - MCC
* @author Naman Jindal <nj.itprof@gmail.com>
* @version 1.0
***************************************************/

var ONLY_ACTIVE_CAMPAIGNS = true;  // true to run script only on active campaigns, false to run on all campaigns

var MASTER_DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
var DASHBOARD_URLS_TAB = 'Dashboard Urls';

function main() {
  
  var masterSS = SpreadsheetApp.openByUrl(MASTER_DASHBOARD_URL);
  var data = masterSS.getSheetByName(DASHBOARD_URLS_TAB).getDataRange().getValues();
  data.shift();
  
  var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy H:mm:ss'));
  var index = NOW.getHours() - 3;
  
  //index = 2;
  if(!data[index]) { return; }
  
  //var ACCOUNT_INPUTS = readAccountSettings();
  
  var ACCOUNT_LABEL = data[index][0];
  var EMAIL = data[index][3];
  
  if(ACCOUNT_LABEL == 'New Business') { return; }
  
  Logger.log('Running for: ' + ACCOUNT_LABEL);
  
  MccApp.accounts()
  .withCondition('LabelNames CONTAINS "' + ACCOUNT_LABEL + '"')
 // .withCondition('Name = "Verex - Volvo"')
  .withCondition('LabelNames DOES_NOT_CONTAIN "Not Live"')
  .executeInParallel('runScript', 'compileResults', EMAIL);
}

// Compile results from all accounts and send email
function compileResults(results) {
  Logger.log('Compiling Results');
  
  var EMAIL = '';
  var adRows = [['Account Name', '# Disapproved Ads']];
  var callRows = [['Account Name', 'Phone Number', 'Device']];
  var calloutRows = [['Account Name', 'Text', 'Device']];
  var slRows = [['Account Name', 'Link Text', 'Link Url']];
  for(var i in results) {
    if(!results[i].getReturnValue()) { continue; }
    var res = JSON.parse(results[i].getReturnValue());
    
    if(res.adRows.length) {
      adRows = adRows.concat(res.adRows);
    }
    
    
    if(res.callRows.length) {
      callRows = callRows.concat(res.callRows);
    }
    
    if(res.slRows.length) {
      slRows = slRows.concat(res.slRows);
    }
    
    if(res.callRows.length) {
      calloutRows = calloutRows.concat(res.calloutRows);
    }
    
    EMAIL = res.EMAIL;
  }

  if(adRows.length == 1 && callRows.length === 1 && slRows.length === 1 && calloutRows.length === 1) { return; }
  
  var htmlBody = '<html><head></head><body>';    
  htmlBody += 'Hi,<br><br>Please find below the Extensions which have been disapproved in your AdWords Accounts:' ;
  
  if(adRows.length > 1) {
    htmlBody += '<br><br><b>Ads (Labeled as Disapproved Ad in account):</b>';
    htmlBody += buildHTMLTable(adRows);    
  }
  
  if(callRows.length > 1) {
    htmlBody += '<br><br><b>Phone Numbers:</b>';
    htmlBody += buildHTMLTable(callRows);    
  }
  
  if(slRows.length > 1) {
    htmlBody += '<br><br><b>Sitelinks:</b>';
    htmlBody += buildHTMLTable(slRows);    
  }
  
  if(calloutRows.length > 1) {
    htmlBody += '<br><br><b>Call Outs:</b>';
    htmlBody += buildHTMLTable(calloutRows);    
  }
  
  htmlBody += '<br><br>Thanks</body></html>';
  var options = { 
    'htmlBody' : htmlBody,
  };
  
  MailApp.sendEmail(EMAIL, 'Disapproved Ads & Extensions (Call/Sitelinks/Callouts) Report', '', options);
}

// Look for disapproved call extensions in the account
function runScript(EMAIL) {
  var accName = AdWordsApp.currentAccount().getName();
  
  var DATE_RANGE = '20000101,'+getAdWordsFormattedDate(0, 'yyyyMMdd');
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['AccountDescriptiveName','FeedItemId','CampaignName','ValidationDetails'];
  var report = 'PLACEHOLDER_FEED_ITEM_REPORT';
  
  
  var query = ['select',cols.join(','),'from',report,
               'where PlaceholderType = 2',
               ONLY_ACTIVE_CAMPAIGNS ? 'and CampaignStatus = ENABLED' : '',
               'during',DATE_RANGE].join(' ');
  
  var ids = [];
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.ValidationDetails.toLowerCase().indexOf('disapprove') < 0) { continue; }
    ids.push(row.FeedItemId);
  }
  
  var callRows = [];  
  if(ids.length) {
    var iter = AdWordsApp.extensions().phoneNumbers().withIds(ids).get();
    while(iter.hasNext()) {
      var ph = iter.next();
      callRows.push([accName, ph.getPhoneNumber(), ph.isMobilePreferred() ? 'Mobile' : 'All']);
    }
  }
  
  var query = ['select',cols.join(','),'from',report,
               'where PlaceholderType = 1',
               ONLY_ACTIVE_CAMPAIGNS ? 'and CampaignStatus = ENABLED' : '',
               'during',DATE_RANGE].join(' ');
  
  var ids = [];
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.ValidationDetails.toLowerCase().indexOf('disapprove') < 0) { continue; }
    ids.push(row.FeedItemId);
  }
  
  var slRows = [];  
  if(ids.length) {
    var iter = AdWordsApp.extensions().sitelinks().withIds(ids).get();
    while(iter.hasNext()) {
      var sl = iter.next();
      slRows.push([accName, sl.getLinkText(), sl.urls().getFinalUrl()]);
    }
  }
  
  
  var query = ['select',cols.join(','),'from',report,
               'where PlaceholderType = 17',
               ONLY_ACTIVE_CAMPAIGNS ? 'and CampaignStatus = ENABLED' : '',
               'during',DATE_RANGE].join(' ');
  
  var ids = [];
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.ValidationDetails.toLowerCase().indexOf('disapprove') < 0) { continue; }
    ids.push(row.FeedItemId);
  }
  
  var calloutRows = [];  
  if(ids.length == 0) {
    var iter = AdWordsApp.extensions().callouts().withIds(ids).get();
    while(iter.hasNext()) {
      var co = iter.next();
      calloutRows.push([accName, co.getText(), co.isMobilePreferred() ? 'Mobile' : 'All']);
    }
  }
  
  if(!AdWordsApp.labels().withCondition('Name = "Disapproved Ad"').get().hasNext()) {
    AdWordsApp.createLabel('Disapproved Ad') 
  }
  
  var adRows = [];
  var iter = AdWordsApp.ads()
  .withCondition('Status = ENABLED')
  .withCondition('AdGroupStatus = ENABLED')
  .withCondition('CampaignStatus = ENABLED')
  .withCondition('ApprovalStatus = DISAPPROVED')
  .get();
  
  if(iter.totalNumEntities() > 0) {
    adRows.push([accName, iter.totalNumEntities()]);
  }
  
  while(iter.hasNext()) {
   var ad = iter.next();
    ad.applyLabel('Disapproved Ad')
  }
  
  return JSON.stringify({ 'adRows': adRows, 'callRows': callRows, 'slRows': slRows, 'calloutRows': calloutRows, 'EMAIL': EMAIL });
}

// Create HTML formatted Email
function buildHTMLTable(DATA) {
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
  
  var header = DATA.shift();
  for(var k in header) {
    table.addHeaderColumn(header[k]);
  }
  
  for(var k in DATA) {
    table.newRow();
    for(var j in DATA[k]) {
      table.addCell(DATA[k][j]);
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

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

