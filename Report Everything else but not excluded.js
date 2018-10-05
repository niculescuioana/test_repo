var EMAIL = 'neeraj@pushgroup.co.uk,sandeep@pushgroup.co.uk,ian@pushgroup.co.uk,naman@pushgroup.co.uk';

function main() {
  MccApp.accounts().withCondition('LabelNames CONTAINS "E-Commerce"').executeInParallel('runScript','compile');
}

function compile(results) {
  var out = []
  for(var i in results) {
    if(!results[i].getReturnValue()) { continue; }
    var rows = JSON.parse(results[i].getReturnValue());
    out = out.concat(rows);
  }
  
  if(!out.length) { return; }
  
  sendEmail(out);
}

function sendEmail(data) {
  var htmlBody = '<html><head></head><body>';    
  htmlBody += 'Please find below the list of shopping Campaigns which have Everything else products which are not excluded.<br>' ;
  htmlBody += buildHTMLTable(data);  
  htmlBody += '<br></body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  
  Logger.log('Sending Report via Email');
  MailApp.sendEmail(EMAIL, 'Shopping Campaigns with Everything else products not excluded', '', options); 
}

function runScript() {
  var OPTIONS = { includeZeroImpressions : true };
  var cols = ['CampaignId', 'CampaignName', 'AdGroupId', 'AdGroupName',
              'Id', 'ProductGroup', 'IsNegative', 'CpcBid'];
  var report = 'PRODUCT_PARTITION_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               //'and CampaignName = "GS - Tisserand"',
               'and ProductGroup = "* / brand = *"',
               'during','YESTERDAY'].join(' ');
  
  var results = {};
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    if(row.CpcBid !== 'Excluded') {
     results[row.CampaignName] = 1; 
    }
  }
  
  var accName = AdWordsApp.currentAccount().getName();
  var out = [];
  for(var camp in results) {
    out.push([accName, camp]);
  }
  
  if(!out.length) {
    return '';  
  }
  
  return JSON.stringify(out);
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
  table.addHeaderColumn('Account Name');
  table.addHeaderColumn('Campaign Name');
  
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