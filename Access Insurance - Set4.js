/*************************************************
* Master Reference Script - Set 4
* @author Naman Jindal <naman@pushgroup.co.uk>
* @version 1.0
***************************************************/

// Url of the Dashboard
// Pick it from Master Dashboard (Column B) -> https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=596395744
var scriptSheet = 'https://docs.google.com/spreadsheets/d/1CZO2ZehKVHqikT5GyORNtxGtIuCVyYkLsHBSgK_IWN8/edit';

// Label for the dashboard (Case Sensitive)
var LABEL = 'Access Insurance';


// Add Email of the new Manager here
var MASTER_EMAIL_LIST = [
  'ricky@pushgroup.co.uk','access_insurance@adsanalyser.com'
];

// Enter the MCC ID of the Client
var SUB_MCC_ID = '';



/********************************** DO NOT EDIT BELOW THIS LINE **********************************/


var masterConfigSheet = 'https://docs.google.com/spreadsheets/d/1Zs0DT7-y5xezWGOHS-eh9myvw-nQQnsFs0ITZpQz26I/edit';
var configSheet = 'Set 4';

var DOES_NOT_CONTAIN_LABEL = 'Not Live';

var SCRIPTS_FOLDER_ID = '0B48ewrnCIsYAMHhfZTNONWFLZnM';

var MASTER_TEMPLATE = 'https://docs.google.com/spreadsheets/d/1y6bZg2sNw_WLMKM80urWLmgAQZ8-bAv1DVu0a-c02JE/edit';
var EMAIL_UTIL_URL = 'https://docs.google.com/spreadsheets/d/1VfKhDpiFeMFPjMBiYT5p4vvimZ-wPG1DgtsZbMVRV7I/edit#gid=0';
var EMAIL_UTIL_ALERT_TAB_NAME = 'Alert Tasks';
var EMAIL_UTIL_TASK_TAB_NAME = 'Tasks';

var ANALYTICS_STATS_REPORT = 'https://docs.google.com/spreadsheets/d/1H3onoG-Pwi6f1GXWqyD2d2MOw3W7E9h87VYuF9Lm-FI/edit';
var ANALYTICS_TAB_NAME = 'Analytics Report';

var BING_URL = 'https://docs.google.com/spreadsheets/d/1Fpyuykm5xeXnJrS6C_5VhWFUX049vjANGEGyiUShp9I/edit';


var SET_COLUMN_INDEX = 0;

//Mock authorizations
// UrlFetchApp.fetch(url); DriveApp.addFile(child); SpreadsheetApp.create(name);
// MailApp.sendEmail(message); XmlService.parse(xml);
//  Analytics.Management.Accounts.list(); SpreadhsheetApp


//Main function - Program starts execution from here
function main() {
  var input = SpreadsheetApp.openByUrl(masterConfigSheet).getSheetByName(configSheet).getDataRange().getValues();
  var accounts = MccApp.accounts();
  
  if(LABEL.length > 0) {
    accounts = MccApp.accounts().withCondition('LabelNames CONTAINS "'+LABEL+'"');
  } else {
    accounts = MccApp.accounts();
  }
  
  if(DOES_NOT_CONTAIN_LABEL.length > 0) {
    accounts = accounts.withCondition('LabelNames DOES_NOT_CONTAIN "'+DOES_NOT_CONTAIN_LABEL+'"');
  }
  
  accounts
  //.withCondition('Name != "Model Guides"')
  .withLimit(50).executeInParallel('runScript', 'compile', JSON.stringify(input));
}

function compile() { }

function runScript(input) {
  if(AdWordsApp.currentAccount().getName() == '') {
    return ''; 
  }
  
  info('Execution Begins');
  var accName = AdWordsApp.currentAccount().getName();
  var accountTZ = AdWordsApp.currentAccount().getTimeZone();
  
  var masterSpreadsheet = SpreadsheetApp.openByUrl(scriptSheet);
  
  var GENERAL_INPUTS = {};
  var generalInputsTabData = masterSpreadsheet.getSheetByName('Report Links').getDataRange().getValues();
  generalInputsTabData.shift();
  
  for(var k in generalInputsTabData) {
    if(!generalInputsTabData[k][2]) { continue; }
    GENERAL_INPUTS[generalInputsTabData[k][2]] = generalInputsTabData[k][1];
  }
  
  var ACCOUNT_INPUTS = {};
  var commonInputsTabData = masterSpreadsheet.getSheetByName('Account Inputs').getDataRange().getValues();
  commonInputsTabData.shift();
  
  var commonInputsHeader = commonInputsTabData.shift();
  for(var k in commonInputsTabData) {
    if(commonInputsTabData[k][0] != accName) { continue; }
    for(var j in commonInputsTabData[k]) {
      ACCOUNT_INPUTS[commonInputsHeader[j]] = commonInputsTabData[k][j];
    }  
  }
  
  var configData = JSON.parse(input);
  
  var EXECUTION_ERRORS = [];
  var days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  
  //Get date as per account Time zone
  var now = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), "MMM dd,yyyy HH:mm:ss"));
  var dateString = getDateString(); //Date in "MM/dd/yyyy" format
  var yyyy = now.getYear();  
  var mm = now.getMonth() + 1;
  var dd = now.getDate();
  var hh = now.getHours();
  var day = days[now.getDay()];	
  
  //Iterate through inputs and execute scripts
  for(var configIndex in configData) {
    if(configIndex < 2) { continue; } //Ignore first 2 rows
    var [description, location, classname, spreadsheet_url, sheetName,
         schedule, time, dayOfMonth, dayOfWeek] = configData[configIndex];
    
    var status = true;
    //Check if script should run or not
    var scheduleDetails = [schedule, time, dayOfMonth,  dayOfWeek];
    status = checkExecutionStatus(scheduleDetails, hh, dd, day);
    
    if(!location || !status) { continue; } //Skip if script location is not specified or script is not scheduled to run now
    
    var rowNum = 0;
    if(sheetName != '') {
      var inputSheet = masterSpreadsheet.getSheetByName(sheetName);
      if(!inputSheet) { continue; }
      if(inputSheet != undefined){
        rowNum = getAccountRowNum(inputSheet, accName); //Get rownum for the account in the input sheet
        if(rowNum == 0) { continue; }
      }
    }

    var attempt = 1, scriptFile, scriptText;
    info(now + ': Running "'+description+'" from location: '+location + ":" + rowNum);
    while(attempt < 10){
      try{
        scriptFile = getFile(location); //Get script file from drive
        scriptText = scriptFile.getBlob().getDataAsString();
        break;
      } catch(e) {
        info(e + ': Retrying ' + attempt + ': ' + location);
        Utilities.sleep(Math.pow(2,attempt)*750 + Math.round(Math.random()*750));
        attempt++;
      }
    }	
    
    if(attempt == 10 || !scriptText) { 
      info('Failed to Load: ' + location); 
      continue; 
    }
    
    try {
      eval(scriptText); //Check for syntax errors		
      var obj = 'new '+classname+'();';
      var script = eval(obj);
      script.main(); //Execute script
      SpreadsheetApp.flush();      
    }  catch(ex){
      var err = ex.constructor('Error in Script: ' + ex.message);
      err.lineNumber = ex.lineNumber - err.lineNumber;
      //info(description + ' - ' + ex + ' (' + err.lineNumber + ')');
      EXECUTION_ERRORS.push(description + ' - ' + ex + ' (' + err.lineNumber + ')');
    }
    
    info('****************************** End of script ******************************');    
  }   
  
  info('Execution Ends');
  
  if(EXECUTION_ERRORS.length > 0) {
    throw AdWordsApp.currentAccount().getName() + ' - One or more Script failed with following reasons: \n' + EXECUTION_ERRORS.join('\n');
  } 
  
  return '';
}


//Today's date in "MM/dd/yyyy" format
function getDateString(){
  return Utilities.formatDate((new Date()), AdWordsApp.currentAccount().getTimeZone(), "MMM dd, yyyy HH:mm:ss");
}

/**
* Get AdWords Formatted date for n days back
* @param {int} d - Numer of days to go back for start/end date
* @return {String} - Formatted date yyyyMMdd
**/
function getAdWordsFormattedDate(d, format){
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
}

/**
* Check scripts schedule against current time
* @param {Array} scheduleDetails - Scheduling Details Array
* @param {int} hh - Current Hour Of Day
* @param {int}dd - Current date
* @param {String} day - Current Day of week
* @return {Boolean} Status - True if script should run, False if not
**/
function checkExecutionStatus(scheduleDetails, hh, dd, day){
  
  if(scheduleDetails[0] == 'Hourly'){
    return true;
  }else if(scheduleDetails[0] == 'Daily' && scheduleDetails[1] == hh){
    return true;
  } else if(scheduleDetails[0] == 'Weekly' && scheduleDetails[3] == day && scheduleDetails[1] == hh){
    return true;
  } else if(scheduleDetails[0] == 'Monthly' && scheduleDetails[2] == dd && (scheduleDetails[1] == '' || scheduleDetails[1] == hh)){
    return true;
  } else if(scheduleDetails[0] == 'Multiple') {
    var hours = scheduleDetails[1].toString().split(',');
    if(hours.indexOf(hh.toString()) > -1) {
      return true; 
    }
  }
  
  return false;	
}


//This function gets the file from GDrive
function getFile(loc) {
  var folder = DriveApp.getFolderById(SCRIPTS_FOLDER_ID);
  if(folder.getFilesByName(loc).hasNext()) {
    return folder.getFilesByName(loc).next();
  } else {
    return null;
  }
}

//This function finds the folder for the file and creates folders if needed
function getFolder(folderPath) {
  var folder = DriveApp.getRootFolder();
  if(folderPath) {
    var pathArray = folderPath.split('/');
    for(var i in pathArray) {
      if(i == pathArray.length - 1) { break; }
      var folderName = pathArray[i];
      if(folder.getFoldersByName(folderName).hasNext()) {
        folder = folder.getFoldersByName(folderName).next();
      }
    }
  }
  return folder;
}

/**
* Check scripts schedule against current time
* @param {Object} sheet - Input sheet
* @param {String} accName - Name of the account
* @return {int} rowNum - Row Num for the account
**/
function getAccountRowNum(sheet,accName){	
  
  var lastRow = sheet.getLastRow();
  var customerName = sheet.getRange("A2:A"+lastRow).getValues();
  
  for(var i = 0; i < customerName.length; i++) {
    if(customerName[i][0] == accName) {
      return (i + 2);
    }	
  }
  
  sheet.getRange(lastRow+1,1).setValue(accName);
  
  info("The Account was not found in the spreadsheet!");
  return 0;
  
}

function info(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + ' - ' + msg);
}

function readCurrencyExchangeRates() {
  var CURRENCY_EXCHANGE_URL = 'https://docs.google.com/spreadsheets/d/1F3bjn411jR3aYEJpLNAeCdqobJYlMQRUlK-6KVcdjCE/edit';
  var CURRENCY_EXCHANGE_TAB_NAME = 'Currency Exchange';
  
  var map = {};
  var data = SpreadsheetApp.openByUrl(CURRENCY_EXCHANGE_URL).getSheetByName(CURRENCY_EXCHANGE_TAB_NAME).getDataRange().getValues();
  data.shift();
  
  for(var k in data) {
    if(!data[k][0]) { continue; }
    map[data[k][0]] = data[k][1]; 
  }
  
  return map;
}

function round(num,n) {    
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

function addToFolder(FOLDER_ID, fileName) {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  
  var fileIter = DriveApp.getRootFolder().searchFiles("title contains '" + fileName + "'");
  while(fileIter.hasNext()){
    var file = fileIter.next();
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }     
}


function deleteExtraRowsCols(sheet, rowFlag, rowOffSet, colFlag, colOffset) {
  if(colFlag) {
    if((sheet.getMaxColumns() - sheet.getLastColumn()) > colOffset) {
      sheet.deleteColumns(sheet.getLastColumn()+(colOffset+1), sheet.getMaxColumns() - sheet.getLastColumn() - colOffset);
    }
  }
  
  if(rowFlag) {
    if((sheet.getMaxRows() - sheet.getLastRow()) > rowOffSet) {
      sheet.deleteRows(sheet.getLastRow()+(rowOffSet+1), sheet.getMaxRows() - sheet.getLastRow() - rowOffSet);
    }
  }
}