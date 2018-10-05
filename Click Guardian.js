
function temp() {
  
    var TEMPLATE = 'https://docs.google.com/spreadsheets/d/1m_aXolaMX0fhGqNdrb0bxwAvOVm3kwlY9wRLwcRB9sw/edit';
    var temp = SpreadsheetApp.openByUrl(TEMPLATE)
    var org = temp.getSheetByName('Weekly Trends');
    var map = {};
    var scriptSheet = 'https://docs.google.com/spreadsheets/d/1ChUlFWpRxsuBFRCLItV_ZCCFczQ7ekKaqRhzsp0tgYg/edit';
    var masterSpreadsheet = SpreadsheetApp.openByUrl(scriptSheet);
    var commonInputsTab = masterSpreadsheet.getSheetByName('Account Inputs');
    var commonInputsTabData = commonInputsTab.getDataRange().getValues();
    commonInputsTabData.shift();
    
    var skip = true;
    var commonInputsHeader = commonInputsTabData.shift();
    var idx = commonInputsHeader.indexOf('REPORT_ID');
    for(var k in commonInputsTabData) {
      if(commonInputsTabData[k][0] == '652-829-0610') { skip = false; }
      if(skip) { continue; }
      /*
      var ss = SpreadsheetApp.openById(commonInputsTabData[k][idx]);
      var sheet = ss.getSheetByName('Weekly Trends');
      var nt = org.copyTo(ss);
      var data = sheet.getDataRange().getValues();
      nt.getRange(1,1,data.length,data[0].length).setValues(data);
      ss.deleteSheet(sheet);
      nt.setName('Weekly Trends');
      
      ss.setActiveSheet(nt);
      ss.moveActiveSheet(4);*/
      
      //Utilities.sleep(100);
      continue;
      //ss.getSheetByName('Duplicate Keywords').getRange('D:L').setHorizontalAlignment('center');
      
      //continue;
      
      var sheets = ss.getSheets();
      for(var z in sheets) {
        var name = sheets[z].getName();
        if(name == 'Free PPC Audit') { continue; }
       
        //sheets[z].getRange(1,1,sheets[z].getMaxRows(),sheets[z].getMaxColumns()).setFontFamily('Poppins');
        /*
        if(org) {
          var nt = org.copyTo(ss);
          var data = sheets[z].getDataRange().getValues();
          nt.getRange(1,1,data.length,data[0].length).setValues(data);
          ss.deleteSheet(sheets[z]);
          nt.setName(name);
        } else {
          ss.deleteSheet(sheets[z]);
        }*/
      }
      
      //break;
    }
  }
  
  //Main function - Program starts execution from here
  function main() {
    //temp();
    //return;
    
    var EMAILS = [
      'analytics@pushgroup.co.uk', 'naman@pushgroup.co.uk', 'charlie@pushgroup.co.uk', 'ricky@pushgroup.co.uk', 'adwords@pushgroup.co.uk',
      'adwords@clickguardian.co.uk', 'clickfraud@clickbuffer.com' 
    ];
    var TEMPLATE = 'https://docs.google.com/spreadsheets/d/1m_aXolaMX0fhGqNdrb0bxwAvOVm3kwlY9wRLwcRB9sw/edit';
    var map = {};
    var scriptSheet = 'https://docs.google.com/spreadsheets/d/1ChUlFWpRxsuBFRCLItV_ZCCFczQ7ekKaqRhzsp0tgYg/edit';
    var masterSpreadsheet = SpreadsheetApp.openByUrl(scriptSheet);
    var commonInputsTab = masterSpreadsheet.getSheetByName('Account Inputs');
    var commonInputsTabData = commonInputsTab.getDataRange().getValues();
    commonInputsTabData.shift();
    
    var commonInputsHeader = commonInputsTabData.shift();
    var idx = commonInputsHeader.indexOf('REPORT_ID');
    for(var k in commonInputsTabData) {
      map[commonInputsTabData[k][0]] = parseInt(k,10)+3;
      if(!commonInputsTabData[k][idx]) {
        var ss = SpreadsheetApp.openByUrl(TEMPLATE).copy(commonInputsTabData[k][1] + ' - Audit');
        ss.addEditors(EMAILS);
        var id = ss.getId(), url = ss.getUrl();
        commonInputsTab.getRange(parseInt(k,10)+3, idx, 1, 2).setValues([[url, id]]);
      }
    }
  
    var out = [];
    //var LABEL = '';
    var accounts = MccApp.accounts()
    .withCondition("ManagerCustomerId IN ['470-648-3435']")
    .withCondition('Impressions > 0')
    .forDateRange('LAST_14_DAYS')
    .get();
    
    while(accounts.hasNext()) {
      MccApp.select(accounts.next()); 
      if(!map[AdWordsApp.currentAccount().getCustomerId()]) { 
        var ss = SpreadsheetApp.openByUrl(TEMPLATE).copy(AdWordsApp.currentAccount().getName() + ' - Audit');
        ss.addEditors(EMAILS);      
        var id = ss.getId(), url = ss.getUrl();
        out.push([AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName(), '', '', url, id]);
        continue; 
      }
      
      delete map[AdWordsApp.currentAccount().getCustomerId()];
    }
    
    if(out.length) {
      commonInputsTab.getRange(commonInputsTab.getLastRow()+1, 1, out.length, out[0].length).setValues(out);
    }
    
    for(var id in map) {
      var row = map[id];
      commonInputsTab.getRange(row, 3).setValue('No');
    }
  }