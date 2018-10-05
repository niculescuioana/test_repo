/******************************************
* Adwords Automated Account Setup using Bulk Upload
* Version 1.0
* Created By: Naman Jindal (nj.itprof@gmail.com)
******************************************/

/************************************************* Configurable Inputs begin ************************************************************/
var SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/19-g4cBNxKAh_PhK8nJklWYAInGN2c9TPScdNxwduqmg/edit#gid=1620253858';
var scriptFileName = 'Scripts/AdParamAutoSpecific.js'; //Location in Google Drive
var doNotTouchLabel = 'Do Not Touch';

/*************************************************** Configurable Inputs end **********************************************************/

function main(){  
  
  //Mock authorizations
  // UrlFetchApp.fetch(url); DriveApp.addFile(child); SpreadsheetApp.create(name);
  // MailApp.sendEmail(message); XmlService.parse(xml);
  // Analytics.Management.Accounts.list();  
  
  var ids = collectAccountsToRun();
  if(ids.length == 0) { Logger.log('No account to run'); }
  MccApp.accounts().withIds(ids).executeInParallel('runScript');
  
}

function collectAccountsToRun() {
  var ids = []
  var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName('Custom Settings').getDataRange().getValues();
  data.shift();
  for(var k in data) {
    if(!data[k][1]) { continue; }
    ids.push(data[k][1]);
  }
  return ids;
}

function runScript() {
  var SETTINGS = initializeScript();
  if(SETTINGS.STATUS == 'OFF') {
    Logger.log('Not Allowed To Run Now');
   // return;
  }
  var code = getCodeFromDoc(scriptFileName);
  
  
  try {
    eval(code);
    var script = eval('new remoteScript();');
    script.main();
  } catch(e){
    var err = e.constructor('Error in Script: ' + e.message);
    err.lineNumber = e.lineNumber - err.lineNumber;
    Logger.log("err line number "+err.lineNumber);
    throw err;
  }
  
  function initializeScript() {
    var HEADER = ['ACC_NAME','ACC_ID','STATUS','FEED_TYPE','FEED_URL','TAB_NAME',
                  'LOCATION_TAB_NAME', 'TEMPLATE_FILE_LOCATION','SYNC',
                  'ADPARAMS','VARIATION_LIST_PRE','VARIATION_LIST_POST'];
    var data = SpreadsheetApp.openByUrl(SETTINGS_URL).getSheetByName('Custom Settings').getDataRange().getValues();
    data.shift();
    var INPUTS = [];
    
    var accId = AdWordsApp.currentAccount().getCustomerId();
    for(var k in data) {
      if(data[k][1] != accId) { continue; }
      
      var SETTINGS = new Object();
      for(var j in HEADER) {
        SETTINGS[HEADER[j]] = data[k][j];
      }
      
      INPUTS.push(SETTINGS);
    }
    
    var NOW = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss'));
    var hour = NOW.getHours();
    
    var index = hour;
    
    while(index >= INPUTS.length) {
     index -= INPUTS.length; 
    }
    
    if(index < 0) { index = 0; }
    Logger.log(index);
    //index = 5;
    
    return INPUTS[index];
  }
  
  function getCodeFromDoc(filename){
    var file = getFile(filename);
    var code = file.getBlob().getDataAsString();
    return code;
  }
  
  function getFile(loc) {
    var locArray = loc.split('/');
    var folder = getFolder(loc);
    if(folder.getFilesByName(locArray[locArray.length-1]).hasNext()) {
      return folder.getFilesByName(locArray[locArray.length-1]).next();
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
}