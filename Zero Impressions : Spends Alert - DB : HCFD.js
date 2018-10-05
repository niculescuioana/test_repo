/******************************************
* Spends Alert
* @version: 1.0
* @author: Naman Jindal (naman@pushgroup.co.uk)
******************************************/

var IDS = ['424-026-1887', '259-199-2135', '754-561-7282'];
var EMAIL = 'Jonathan@jewellerycave.co.uk,TGardiner@jewellerycave.co.uk,monique@pushgroup.co.uk,aurelia@pushgroup.co.uk';
var DATE_RANGE = 'YESTERDAY';


function main() {
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
  
  if(now.getHours() === 9) {
    sendSpendsReport();
  }
  
  if(now.getHours() > 8 && now.getHours() < 20) {
    sendImpressionsAlert();
  } 
}

function sendImpressionsAlert() {
  var out = [['Customer Id', 'Customer Name']];
  var iter = MccApp.accounts().withIds(IDS).get();
  while(iter.hasNext()) {
   var account = iter.next();
    MccApp.select(account);
    
    var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy HH:mm'));
    var hour = date.getHours() - 4;
    
    var OPTIONS = { includeZeroImpressions : false };
    var cols = ['HourOfDay','Impressions'];
    var report = 'ACCOUNT_PERFORMANCE_REPORT';
    var query = ['select',cols.join(','),'from',report,
                 'where HourOfDay >= '+hour,
                 'during','TODAY'].join(' ');
    
    var reportIter = AdWordsApp.report(query, OPTIONS).rows();
    if(!reportIter.hasNext()) {
      out.push([AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName()]);
    }
  }

  if(out.length < 2) { return; }
  
  var SUBJECT = 'Zero Impressions Alert';
  var htmlBody = '<html><head></head><body>';    
  htmlBody += "Attention! Below Accounts have not recieved any impressions in last 4 hours.<br>";
  htmlBody += buildHTMLTable(out);  
  htmlBody += '<br><br>Thanks</body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  
  MailApp.sendEmail(EMAIL, SUBJECT, '', options);
}

function sendSpendsReport() {
  var out = [['Customer Id', 'Customer Name', 'Spends']];
  var iter = MccApp.accounts().withIds(IDS).get();
  while(iter.hasNext()) {
   var account = iter.next();
    MccApp.select(account);
    
    var spends = AdWordsApp.currentAccount().getStatsFor(DATE_RANGE).getCost();
    out.push([AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName(), spends]);
  }

  if(out.length < 2) { return; }
             
  var SUBJECT = 'Daily Spends Report';
  var htmlBody = '<html><head></head><body>';    
  htmlBody += "Please find below yesterday's spends summary for your accounts<br>";
  htmlBody += buildHTMLTable(out);  
  htmlBody += '<br><br>Thanks</body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  
  MailApp.sendEmail(EMAIL, SUBJECT, '', options);
}

function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}


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
  for(var x in header) {
    table.addHeaderColumn(header[x]);
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