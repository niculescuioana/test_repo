/******************************************
* CPC Alert
* @version: 1.0
* @author: Naman Jindal (nj.itprof@gmail.com)
******************************************/

var EMAIL_TO = 'hs@pushgroup.co.uk,sandeep@pushgroup.co.uk,neeraj@pushgroup.co.uk,ian@pushgroup.co.uk';
var EMAIL_CC = 'ricky@pushgroup.co.uk,charlie@pushgroup.co.uk';
var EMAIL_BCC = 'naman@pushgroup.co.uk';

/*
var EMAIL_TO = 'nj.itprof@gmail.com';
var EMAIL_CC = 'namanjindal12345@gmail.com';
var EMAIL_BCC = 'naman@pushgroup.co.uk';
*/

var EMAIL_UTIL_URL = 'https://docs.google.com/spreadsheets/d/1VfKhDpiFeMFPjMBiYT5p4vvimZ-wPG1DgtsZbMVRV7I/edit#gid=0';
var EMAIL_UTIL_TAB_NAME = 'Client Emails';
var DATE_NOW = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss');


function main() {
  MccApp.accounts().withIds(['309-585-8616']).executeInParallel('runScript');
}

function runScript() {
  var YESTERDAY = getAdWordsFormattedDate(1, 'yyyyMMdd');
  var LAST_90_DAYS = getAdWordsFormattedDate(90, 'yyyyMMdd') + ',' + getAdWordsFormattedDate(1,  'yyyyMMdd');
  
  var cpcMap = {};
  
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['CampaignName','AdGroupName','AdGroupId','AverageCpc'];
  var report = 'ADGROUP_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               'during',LAST_90_DAYS].join(' ');
  
  var iter = AdWordsApp.report(query, OPTIONS).rows();
  while(iter.hasNext()) {
    var row = iter.next();
    cpcMap[row.AdGroupId] = parseFloat(row.AverageCpc.toString().replace(/,/g,''));
  }
  
  var toReport = [];
  var query = ['select',cols.join(','),'from',report,
               'where AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               'during','YESTERDAY'].join(' ');
  
  var iter = AdWordsApp.report(query, OPTIONS).rows();
  while(iter.hasNext()) {
    var row = iter.next();
    if(!cpcMap[row.AdGroupId]) { continue; }
    
    var cpc = parseFloat(row.AverageCpc.toString().replace(/,/g,''));
    if(cpc <= cpcMap[row.AdGroupId]) { continue; }
    
    toReport.push([row.CampaignName, row.AdGroupName, cpc, cpcMap[row.AdGroupId]]);
  }
  
  if(!toReport.length) { return; }
  
  sendEmail(toReport);
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}

function sendEmail(toReport) {
  var YESTERDAY = getAdWordsFormattedDate(1, 'MMM d');
  var LAST_90_DAYS = getAdWordsFormattedDate(90, 'MMM d') + ' - ' + YESTERDAY;
  
  Logger.log('Sending Email');
  var SUB = AdWordsApp.currentAccount().getName() + " - High CPC AdGroups Alert";
  
  var htmlBody = '<html><head></head><body><font face="Trebuchet MS">'
  htmlBody += 'Hi,<br><br>Below is the summary of AdGroups where CPC yesterday was above 90 days average.'
  
  var header = ['Campaign', 'AdGroup', 'Avg Cpc (' + YESTERDAY + ')', 'Avg Cpc (' + LAST_90_DAYS + ')'];
  htmlBody += buildReport(toReport, header, '#6678b1');
  
  htmlBody += '<br><br>Thanks</font</body></html>';
  var options = { 
    htmlBody : htmlBody
  };
  
 // MailApp.sendEmail(EMAIL_TO, SUB, '', options);
  
 // return;
  
  var EMAIL_UTIL_TAB = SpreadsheetApp.openByUrl(EMAIL_UTIL_URL).getSheetByName(EMAIL_UTIL_TAB_NAME);
  EMAIL_UTIL_TAB.appendRow([EMAIL_TO, EMAIL_CC, EMAIL_BCC, SUB, '', htmlBody, DATE_NOW]);
}


function buildReport(toReport, header, color) {
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
                        'color: ' + color +';',
                        'padding: 10px 8px;',
                        'border-bottom: 2px solid ' + color].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  
  for(var j in header) {
    table.addHeaderColumn(header[j]);
  }
  
  for(var k in toReport) {
    table.newRow();
    for(var j in toReport[k]) {
      table.addCell(toReport[k][j]);	
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