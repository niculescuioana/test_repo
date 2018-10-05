var MANAGER = 'Yannis';  
var DOES_NOT_CONTAIN_LABEL = 'Not Live';

function main() {
  MccApp.accounts()
  .withCondition('LabelNames DOES_NOT_CONTAIN "'+DOES_NOT_CONTAIN_LABEL+'"')
  .withCondition('LabelNames CONTAINS "'+MANAGER+'"')
  .executeInParallel('runScript','compileReport');
  
}

function compileReport(results) {
  var out = [['Account', 'Campaign', 'Impressions', 'Clicks', 'Position']];
  for(var z in results) {
    if(!results[z].getReturnValue()) { continue; }
    var rows = JSON.parse(results[z].getReturnValue());
    out = out.concat(rows);
  }
  
  if(out.length === 1) { return; }
  
  var SUBJECT = 'Branding Campaigns Overview - Yesterday'; 
  
  var htmlBody = '<html><head></head><body>';    
  htmlBody += buildTable(out); 
  htmlBody += '<br>Thanks</body></html>';
  
  MailApp.sendEmail('yannis@pushgroup.co.uk', SUBJECT, '', { 'htmlBody': htmlBody });
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
    table.newRow();
    for(var j in reportData[k]){
      table.addCell(reportData[k][j]);
    }  
  }
  return table.toString();
}



function runScript() {
  
  if(!AdWordsApp.labels().withCondition('Name = "01. Branded"').get().hasNext()) { return ''; }

  var toReport = [];
  var iter = AdWordsApp.campaigns()
  .withCondition('LabelNames CONTAINS_ANY ["01. Branded"]')
  .withCondition('Impressions > 0')
  .forDateRange('YESTERDAY')
  .get();
  
  while(iter.hasNext()) {
    var camp = iter.next();
    var stats = camp.getStatsFor('YESTERDAY');
    toReport.push([AdWordsApp.currentAccount().getName(), camp.getName(),
                   stats.getImpressions(), stats.getClicks(), stats.getAveragePosition()]);
  }
  
  if(!toReport.length) { return ''; }
  
  
  return JSON.stringify(toReport)
  
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

