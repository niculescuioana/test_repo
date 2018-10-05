var URL = 'https://docs.google.com/spreadsheets/d/1QBiZCvRIQ0aqqYQxcuLWlSkSldsHxS4Pe82yMV_YlaQ/edit';
var N = 1;

var EMAILS = [
  'neeraj@pushgroup.co.uk',
  'ian@pushgroup.co.uk',
  'jay@pushgroup.co.uk ',
  'sunny@pushgroup.co.uk ',
  'hs@pushgroup.co.uk',
  'ricky@pushgroup.co.uk',
  'charlie@pushgroup.co.uk',
  'jane@personalisedgiftsshop.co.uk',
  'alex@personalisedgiftsshop.co.uk'
].join(',');

var PREVIEW_EMAILS = [
  'neeraj@pushgroup.co.uk',
  'ian@pushgroup.co.uk',
  'naman@pushgroup.co.uk'
].join(',');


var ALERT_EMAILS = [
  'neeraj@pushgroup.co.uk',
  'ian@pushgroup.co.uk',
  'jay@pushgroup.co.uk ',
  'sunny@pushgroup.co.uk ',
  'ricky@pushgroup.co.uk',
  //'naman@pushgroup.co.uk'
  'charlie@pushgroup.co.uk'
  
];
                   
function main() {
  var date = new Date(getAdWordsFormattedDate(N, 'MMM d, yyyy HH:mm'));
  var dd = date.getDate();
  date.setDate(1);
  date.setMonth(date.getMonth()+1);
  date.setDate(0);
  
  var daysLeft = date.getDate() + 1 - dd;
  var YESTERDAY = getAdWordsFormattedDate(N, 'yyyy-MM-dd');
  var parts = YESTERDAY.split('-')
  parts[0] = parseInt(parts[0],10)-1;
  var YEST_LY = parts.join('-');
  
  var filters = ['ga:medium==cpc;ga:source==google'];
  var filters_1 = ['mcf:medium==cpc;mcf:source==google'];
  var googleStats = {
    'PGS': {
      'yest': getStatsForDate(17367336, YESTERDAY, filters, filters_1),
      'yestLy': getStatsForDate(17367336, YEST_LY, filters, filters_1)
    },
    'PWG': {
      'yest': getStatsForDate(43299184, YESTERDAY, filters, filters_1),
      'yestLy': getStatsForDate(43299184, YEST_LY, filters, filters_1)
    }/*,
    'Wells': {
      'yest': getStatsForDate(137608301, YESTERDAY, filters, filters_1),
      'yestLy': getStatsForDate(137608301, YEST_LY, filters, filters_1)
    }*/
  }
  
  var filters = ['ga:source=@bing,ga:source=@Bing;ga:medium==cpc,ga:medium==CPC'];
  var filters_1 = ['mcf:source=@bing,mcf:source=@Bing;mcf:medium==cpc,mcf:medium==CPC'];
  var bingStats = {
    'PGS': {
      'yest': getStatsForDate(17367336, YESTERDAY, filters, filters_1),
      'yestLy': getStatsForDate(17367336, YEST_LY, filters, filters_1)
    },
    'PWG': {
      'yest': getStatsForDate(43299184, YESTERDAY, filters, filters_1),
      'yestLy': getStatsForDate(43299184, YEST_LY, filters, filters_1)
    }/*,
    'Wells': {
      'yest': { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0 },
      'yestLy': { 'Cost': 0, 'Conversions': 0, 'ConversionValue': 0 }
    }*/
  }
  
  addBingSpends(bingStats);
  
  var ss = SpreadsheetApp.openByUrl(URL);
  var TM = getAdWordsFormattedDate(N, 'MMMM yyyy');
  var tab = ss.getSheetByName(TM);
  
  if(!tab) {
    var dt = new Date(getAdWordsFormattedDate(N, 'MMM d, yyyy'));
    dt.setMonth(dt.getMonth()-1);
    dt.setHours(12);
    var LM = Utilities.formatDate(dt, 'PST', 'MMMM yyyy');
    tab = setupTab(TM, LM); 
  }
  
  var column = parseInt(getAdWordsFormattedDate(N, 'd'), 10) + 2;
  tab.getRange(3,column).setValue(googleStats['PGS']['yest'].ConversionValue);
  tab.getRange(4,column).setValue(googleStats['PWG']['yest'].ConversionValue);
  tab.getRange(5,column).setValue(bingStats['PGS']['yest'].ConversionValue);
  tab.getRange(6,column).setValue(bingStats['PWG']['yest'].ConversionValue);
  //tab.getRange(7,column).setValue(googleStats['Wells']['yest'].ConversionValue);
  
  
  tab.getRange(11,column).setValue(googleStats['PGS']['yestLy'].ConversionValue + bingStats['PGS']['yestLy'].ConversionValue);
  tab.getRange(12,column).setValue(googleStats['PWG']['yestLy'].ConversionValue + bingStats['PWG']['yestLy'].ConversionValue);
  
  tab.getRange(15,column).setValue(googleStats['PGS']['yestLy'].Cost + bingStats['PGS']['yestLy'].Cost);
  tab.getRange(16,column).setValue(googleStats['PWG']['yestLy'].Cost + bingStats['PWG']['yestLy'].Cost);
  
  tab.getRange(22,column).setValue(googleStats['PGS']['yest'].Conversions + bingStats['PGS']['yest'].Conversions);
  tab.getRange(23,column).setValue(googleStats['PWG']['yest'].Conversions + bingStats['PWG']['yest'].Conversions);
  //tab.getRange(24,column).setValue(googleStats['Wells']['yest'].Conversions);  
  
  tab.getRange(32,column).setValue(googleStats['PGS']['yest'].Cost);
  tab.getRange(33,column).setValue(bingStats['PGS']['yest'].Cost);
  tab.getRange(34,column).setValue(googleStats['PWG']['yest'].Cost);
  tab.getRange(35,column).setValue(bingStats['PWG']['yest'].Cost);
  //tab.getRange(36,column).setValue(googleStats['Wells']['yest'].Cost);
  
  var totalRequired = ((tab.getRange('B29').getValue()-tab.getRange('B9').getValue())+tab.getRange(9,column).getValue())/daysLeft;
  tab.getRange(29, column).setValue(totalRequired);
  
  var pctToTarget = tab.getRange('B9').getValue()/tab.getRange('B29').getValue();
  tab.getRange(30, column).setValue(pctToTarget);
  
  var sheet = ss.getSheetByName('Daily Update');
  sheet.getRange(5,4,5,1).setValues(sheet.getRange(5,2,5,1).getValues());
  sheet.getRange(6,9).setValue(totalRequired);
  sheet.getRange(7,9).setValue(pctToTarget);
  
  
  sheet.getRange(13,2,2,2).setValues([[tab.getRange('B3').getValue(), tab.getRange('B32').getValue()],
                                      [tab.getRange('B5').getValue(), tab.getRange('B33').getValue()]]);
  
  sheet.getRange(13,5,2,2).setValues([[tab.getRange('B4').getValue(), tab.getRange('B34').getValue()],
                                      [tab.getRange('B6').getValue(), tab.getRange('B35').getValue()]]);
  
  //sheet.getRange(13,8,1,2).setValues([[tab.getRange('B7').getValue(), tab.getRange('B36').getValue()]]);
  
  
  
  sheet.getRange(19,2,2,2).setValues([[googleStats['PGS']['yest'].ConversionValue, googleStats['PGS']['yest'].Cost],
                                      [bingStats['PGS']['yest'].ConversionValue,bingStats['PGS']['yest'].Cost]]);
  
  sheet.getRange(19,5,2,2).setValues([[googleStats['PWG']['yest'].ConversionValue, googleStats['PWG']['yest'].Cost],
                                      [bingStats['PWG']['yest'].ConversionValue,bingStats['PWG']['yest'].Cost]]);
  
 //sheet.getRange(19,8,2,2).setValues([[googleStats['Wells']['yest'].ConversionValue, googleStats['Wells']['yest'].Cost],
   //                                   [bingStats['Wells']['yest'].ConversionValue,bingStats['Wells']['yest'].Cost]]);
  
  
  
  var day = Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'EEEE')
  if(day == 'Saturday' || day == 'Sunday') { return; }
  
  if(!AdWordsApp.getExecutionInfo().isPreview()) {
    notifyByEmail(EMAILS);
    sendAlertEmail();
  } else {
    notifyByEmail(PREVIEW_EMAILS);
  }
}

function addBingSpends(bingStats) {
  
  var bingName = {
    'PGS': 'pgsgifts',
    'PWG': 'pwggifts'
  }
  
  var statsMap = readStatsFromBing();
  
  var YESTERDAY = getAdWordsFormattedDate(N, 'yyyy-MM-dd');
  var parts = YESTERDAY.split('/')
  parts[2] = parseInt(parts[0],10)-1;
  var YEST_LY = parts.join('/');
  
  if(statsMap[bingName['PGS']] && statsMap[bingName['PGS']][YESTERDAY]) {
    bingStats['PGS']['yest'].Cost = statsMap[bingName['PGS']][YESTERDAY].Cost ;
  }
  
  if(statsMap[bingName['PGS']] && statsMap[bingName['PGS']][YEST_LY]) {
    bingStats['PGS']['yestLy'].Cost = statsMap[bingName['PGS']][YEST_LY].Cost ;
  }
  
  
  if(statsMap[bingName['PWG']] && statsMap[bingName['PWG']][YESTERDAY]) {
    bingStats['PWG']['yest'].Cost = statsMap[bingName['PWG']][YESTERDAY].Cost ;
  }
  
  if(statsMap[bingName['PSG']] && statsMap[bingName['PWG']][YEST_LY]) {
    bingStats['PWG']['yestLy'].Cost = statsMap[bingName['PWG']][YEST_LY].Cost ;
  }
}


function readStatsFromBing() {
  var folder = DriveApp.getFolderById('0BwnikHB3eS37SnU2c3h3Y0wyNGs');
  var statsMap = {};
  
  var data = folder.getFilesByName('Daily Report Stats - TM.csv').next().getBlob().getDataAsString();
  var content = Utilities.parseCsv(data);
  content.shift();
  
  for(var x in content) {
    if(!statsMap[content[x][2]]) {
      statsMap[content[x][2]] = {};
    }
    
    statsMap[content[x][2]][content[x][0]] = {
      'Cost': parseFloat(content[x][5].toString().replace(/,/g,'')),
      'Clicks': parseInt(content[x][4], 10),
      'Conversions': parseInt(content[x][6], 10)
    }
  }
  
  var data = folder.getFilesByName('Daily Report Stats - LM.csv').next().getBlob().getDataAsString();
  var content = Utilities.parseCsv(data);
  content.shift();
  
  for(var x in content) {
    if(!statsMap[content[x][2]]) {
      statsMap[content[x][2]] = {};
    }
    
    statsMap[content[x][2]][content[x][0]] = {
      'Cost': parseFloat(content[x][5].toString().replace(/,/g,'')),
      'Clicks': parseInt(content[x][4], 10),
      'Conversions': parseInt(content[x][6], 10)
    }
  }
  
  var data = folder.getFilesByName('Daily Report Stats - LY.csv').next().getBlob().getDataAsString();
  var content = Utilities.parseCsv(data);
  content.shift();
  
  for(var x in content) {
    if(!statsMap[content[x][2]]) {
      statsMap[content[x][2]] = {};
    }
    
    statsMap[content[x][2]][content[x][0]] = {
      'Cost': parseFloat(content[x][5].toString().replace(/,/g,'')),
      'Clicks': parseInt(content[x][4], 10),
      'Conversions': parseInt(content[x][6], 10)
    }
  }
  
  return statsMap
}

function getStatsForDate(PROFILE_ID, DATE, filters, filters_1) {
  
  var stats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0 
  };
  
  getDataFromAnalytics(PROFILE_ID,stats,DATE,DATE,filters);
  getDataFromMCF(PROFILE_ID,stats,DATE,DATE,filters_1);
  
  stats.CPA = stats.Conversions == 0 ? 0 : round(stats.Cost/stats.Conversions,2);
  stats.ROAS = stats.Cost == 0 ? '0.00%' : round(100*stats.ConversionValue/stats.Cost,2)+'%';
  
  
  return stats;
}


function getDataFromAnalytics(PROFILE_ID,stats,FROM,TO,filters) {
  var optArgs = { 'filters': filters.join(';') };
  var attempts = 3;
  
  // Make a request to the API.
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue",
        optArgs);
      
      break;
    } catch(ex) {
      Logger.log(ex + " ID: " + PROFILE_ID);
      attempts--;
      Utilities.sleep(2000);
    }
  }
  
  var rows = resp.getRows();
  
  for(var k in rows) {
    stats.Cost += parseFloat(rows[k][0]);
    stats.Conversions += parseInt(rows[k][1],10);
    stats.ConversionValue += parseFloat(rows[k][2]);
  }
}

function getDataFromMCF(PROFILE_ID,stats,FROM, TO, filters) {
  
  return;
  
  filters.push('mcf:basicChannelGroupingPath=@Paid Search',
               'mcf:basicChannelGroupingPath!=Paid Search',
               'mcf:conversionType==Transaction');
  
  var optArgs = { 
    'dimensions': 'mcf:basicChannelGroupingPath,mcf:conversionType',
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
    
    stats.Conversions += parseInt(rows[k][2].primitiveValue,10);
    stats.ConversionValue += parseFloat(rows[k][3].primitiveValue);    
  }
}

function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

function setupTab(name, LM) {
  var ss = SpreadsheetApp.openByUrl(URL);
  ss.setActiveSheet(ss.getSheetByName(LM))
  var tab = ss.duplicateActiveSheet();
  tab.setName(name);
  tab.showSheet();
  
  var date = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM d, yyyy HH:mm'));
  date.setHours(12);
  date.setDate(1);
  
  var header_1 = [], header_2 = [];
  
  var month = date.getMonth();
  while(date.getMonth() == month) {
    header_1.push(Utilities.formatDate(date, 'PST', 'EEEE'));
    header_2.push(Utilities.formatDate(date, 'PST', 'd-MMM-yyyy'));
    date.setDate(date.getDate()+1);
  }
  
  tab.getRange(1,3,2,header_1.length).setValues([header_1,header_2]);
  
  tab.getRange('C3:AG7').clearContent();
  tab.getRange('C22:AG24').clearContent();
  tab.getRange('C29:AG36').clearContent();
  
  return tab;
}



function sendAlertEmail() {
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Daily Update');
  
  var toReport = [];
  var target = tab.getRange(5,2).getValue();
  var data = tab.getRange(7,1,2,2).getValues();
  for(var z in data) {
    if(data[z][1] < target) {
      data[z][1] = round_(100*data[z][1], 2) +'%';
      toReport.push(data[z]);
    }  
  }
  
  if(!toReport.length) { return; }
  
  var htmlBody = '<html><head></head><body><font family="Calibri">';    
  htmlBody += 'Hi All,<br><br>ROAS for below listed accounts is below the Target:<br><br>';
  htmlBody += buildSummary(toReport);
  htmlBody += '<br><br>Kind Regards,<br>Push Scripts</font></body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  
  MailApp.sendEmail(ALERT_EMAILS, 'PGS/PWG/5Wells Report - ROAS Below Target', '', {'htmlBody': htmlBody} );
}


function notifyByEmail(LIST) {
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName('Daily Update');
  
  var data = tab.getRange(5,1,4,3).getValues();
  for(var z in data) {
    if(data[z][2] != '-') {
      data[z][2] = round_(100*data[z][2], 2) +'%';
      if(data[z][2].indexOf('-') === -1) { data[z][2] = '+' + data[z][2]; }
    }
    
    data[z][1] = round_(100*data[z][1], 2) +'%';
  }
  
  var data_1 = tab.getRange(6,8,2,3).getValues();
  data_1[0][1] = '£'+round_(data_1[0][1], 2);
  data_1[1][1] = round_(100*data_1[1][1], 2) +'%';
  
  data = data.concat(data_1);
  
  var toReport = tab.getRange(12, 1, 4, 10).getValues();
  var toReport_2 = tab.getRange(18, 1, 4, 10).getValues();
  
  var htmlBody = '<html><head></head><body><font family="Calibri">';    
  htmlBody += 'Hi Alex & Jane,<br><br>Please find a summary below:<br><br>';
  htmlBody += buildSummary(data);
  
  
  htmlBody += '<br><br>MTD Report:<br>';
  htmlBody += buildReport(toReport);
  
  /*htmlBody += '<br><br>Yesterday Report:<br>';
  htmlBody += buildReport(toReport_2);*/
  
  htmlBody += '<br><br>Kind Regards,<br>Push</font></body></html>';
  var options = { 
    htmlBody : htmlBody,
  };
  
  MailApp.sendEmail(LIST, 'Push PGS Daily Performance Update', '', {'htmlBody': htmlBody} );
}

function buildSummary(toReport) {
  var table = new HTMLTable();
  table.setTableStyle(['font-family: "Calibri";',
                       'font-size: 11px;',
                       'background: #fff;',
                       'margin: 45px;',
                       'width: 480px;',
                       'border-collapse: collapse;',
                       'text-align: left'].join(''));
  table.setHeaderStyle(['font-size: 14px;',
                        'font-weight: normal;',
                        'color: #039;',
                        'padding: 10px 8px;',
                        'border-bottom: 3px solid #6678b1'].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  
  table.addHeaderColumn('');
  table.addHeaderColumn('MTD');
  table.addHeaderColumn('Change');
  
  for(var z in toReport) {
    table.newRow();
    for(var y in toReport[z]) {
      table.addCell(toReport[z][y]);	
    }
  }
   
  return table.toString();
}
      

function buildReport(toReport) {
  var table = new HTMLTable();
  table.setTableStyle(['font-family: "Calibri";',
                       'font-size: 11px;',
                       'background: #fff;',
                       'margin: 45px;',
                       'width: 480px;',
                       'border-collapse: collapse;',
                       'text-align: left'].join(''));
  table.setHeaderStyle(['font-size: 14px;',
                        'font-weight: normal;',
                        'color: #039;',
                        'padding: 10px 8px;',
                        'border-bottom: 3px solid #6678b1'].join(''));
  table.setCellStyle(['border-bottom: 1px solid #ccc;',
                      'padding: 4px 6px'].join(''));
  table.addHeaderColumn('',1,2);
  table.addHeaderColumn('PGS',3);
  table.addHeaderColumn('PWG',3);
  table.addHeaderColumn('Overall',3);
  
  var header = toReport.shift();
  for(var z in header) {
    table.addCell(header[z],'font-size: 12px; color: #6aa84f; padding: 10px 8px; border-bottom: 3px solid #d9ead3');	
  }
  
  
  for(var k in toReport) {
    table.newRow();
    for(var j in toReport[k]) {
      var val = toReport[k][j];
      if(j != 0) {
        if(j%3 == 0) {
          val = round_(100*toReport[k][j],2)+'%';
        } else {
          val = '£'+ round_(toReport[k][j],2); 
        }
      }
      table.addCell(val);	
    }
  }
  
  return table.toString();
}



function round_(num,n) {
  if(isNaN(num)) {
    return num; 
  }
  
  return +(Math.round(num + "e+"+n)  + "e-"+n);
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
  
  this.addHeaderColumn = function(text, colspan, rowspan) {
    this.headers.push({ val:text, colspan: (colspan) ? colspan : 1, rowspan: (rowspan) ? rowspan : 1 });
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
  
  function _th(instance,obj) {
    var retVal = '<th scope="col" colspan='+obj.colspan+' rowspan='+obj.rowspan+' ';
    if(instance.headerStyle) {
      retVal += 'style="'+instance.headerStyle+'"';
    }
    retVal += '>'+obj.val+'</th>';
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


function round(num,n) {    
  return +(Math.round(num + "e+"+n)  + "e-"+n);
}