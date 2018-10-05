function main() {
  
  var LABEL_NAME = 'Access Insurance';  // Input Name of Label for new accounts
  var MANAGER_EMAIL = 'access_insurance@adsanalyser.com';  // Input main stakeholder email
  var EXEC_EMAILS = ['access_insurance@adsanalyser.com']; // Input backup stakeholders emails.
  
  
  /********************* No Edits below this please ***************************/
  
  var TEMPLATES = {
    'MANAGEMENT_URL': 'https://docs.google.com/spreadsheets/d/1e7kdFpXPw7vsGGvczfg6Jg8-ZexEQY9msSXBh5PGNb4/edit',
    'HOLIDAY_MANAGER_URL': 'https://docs.google.com/spreadsheets/d/13k68suGmN04kNXxr7ZqHQo5lotNS_sjqYLUG8bt6hJw/edit',
    'CALL_SCHEDULE_MANAGER_URL': 'https://docs.google.com/spreadsheets/d/1-5lC_FO_N7F0Y5FS502IeOVCXykg8bjCTFFmDO5V1uE/edit',
    'CAMPAIGN_LABEL_REPORT_URL': 'https://docs.google.com/spreadsheets/d/102QEj8NWFWa_spabOwXs0Si7aFOuW5nZGUoyOgMEmzg/edit#gid=893920484',
    'ADGROUP_LABEL_REPORT_URL': 'https://docs.google.com/spreadsheets/d/102QEj8NWFWa_spabOwXs0Si7aFOuW5nZGUoyOgMEmzg/edit#gid=893920484',
    'KEYWORD_LABEL_REPORT_URL': 'https://docs.google.com/spreadsheets/d/102QEj8NWFWa_spabOwXs0Si7aFOuW5nZGUoyOgMEmzg/edit#gid=893920484',
    'AD_LABEL_REPORT_URL': 'https://docs.google.com/spreadsheets/d/102QEj8NWFWa_spabOwXs0Si7aFOuW5nZGUoyOgMEmzg/edit#gid=893920484'
  };
  
  var SCRIPTS_FOLDER = '0B51HFuINK5uhaHc4ZGhPQjhTdVU';
  /*if(EXEC_EMAILS.length) {
    DriveApp.getFolderById(SCRIPTS_FOLDER).addEditor(MANAGER_EMAIL).addEditors(EXEC_EMAILS);
  } else {
    DriveApp.getFolderById(SCRIPTS_FOLDER).addEditor(MANAGER_EMAIL);
  }*/
  
  var MASTER_DASHBOARD = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit#gid=1415541725';
  var master_tab = SpreadsheetApp.openByUrl(MASTER_DASHBOARD).getSheetByName('External Dashboard Urls');
  var input = master_tab.getDataRange().getValues();
  input.shift();
  
  var d_map = {};
  for(var z in input) {
    if(input[z][1]) {
      d_map[input[z][0]] = { 'URL': input[z][1], 'FOLDER_ID': input[z][5] };  
    }
  }
  
  var MANAGER_REPORTS_FOLDER_ID = '0B48ewrnCIsYAWnhQNVhack5JZkE';
  
  var DASHBOARD_TEMPLATE = 'https://docs.google.com/spreadsheets/d/1Po2x8jS_BLAv7zkF_IQG4ISB2HTPwlP2cKokir7vaTM/edit';
  var PUSH_TEMPLATES = 'https://docs.google.com/spreadsheets/d/1y6bZg2sNw_WLMKM80urWLmgAQZ8-bAv1DVu0a-c02JE/edit';
  var ss, MANAGER_FOLDER
      
  if(!d_map[LABEL_NAME]) {
    MANAGER_FOLDER = DriveApp.getFolderById(MANAGER_REPORTS_FOLDER_ID).createFolder(LABEL_NAME);
    ss = SpreadsheetApp.openByUrl(DASHBOARD_TEMPLATE).copy('Scripts Dashboard - ' + LABEL_NAME);
    Logger.log('New Dashboard: ' + ss.getUrl());
    Utilities.sleep(1000);
    
    MANAGER_FOLDER.addFile(DriveApp.getFileById(ss.getId()));
    master_tab.appendRow([LABEL_NAME, ss.getUrl(), '', MANAGER_EMAIL, '', MANAGER_FOLDER.getId()]);
  } else {
    ss =  SpreadsheetApp.openByUrl(d_map[LABEL_NAME].URL);
    
    if(d_map[LABEL_NAME].FOLDER_ID) {
      MANAGER_FOLDER = DriveApp.getFolderById(d_map[LABEL_NAME].FOLDER_ID);
    }
  }
  
  var tab = ss.getSheetByName('Report Links');
  tab.getRange('B2').setValue(MANAGER_EMAIL);
  tab.getRange('B3').setValue(EXEC_EMAILS.join(','));
  
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();
  
  for(var z in data) {
    if(data[z][2] && !data[z][1]) {
      var sh;
      if(TEMPLATES[data[z][2]]) {
        sh = SpreadsheetApp.openByUrl(TEMPLATES[data[z][2]]).copy(data[z][0] + ' - ' + LABEL_NAME);
      } else {
        sh = SpreadsheetApp.create(data[z][0] + ' - ' + LABEL_NAME);
      }
      
      tab.getRange(parseInt(z,10) + 4, 2).setValue(sh.getUrl());
      
      if(MANAGER_FOLDER) {
        MANAGER_FOLDER.addFile(DriveApp.getFileById(sh.getId()));
      }
    }
  }
  
}