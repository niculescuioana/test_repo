function main() {
  
    var DASHBOARD_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
    var URL_INFO_TAB = 'Dashboard Urls';
    
    
    var urlData = SpreadsheetApp.openByUrl(DASHBOARD_URL).getSheetByName(URL_INFO_TAB).getDataRange().getValues();
    urlData.shift();
    for(var j in urlData) {
      if(!urlData[j][0]) { continue; }
      //if(urlData[j][0] != 'Yannis') { continue; }
      
      var url = '';
      var data = SpreadsheetApp.openByUrl(urlData[j][1]).getSheetByName('Report Links').getDataRange().getValues();
      for(var x in data) {
        if(data[x][2] === 'HOLIDAY_MANAGER_URL') {
         url =  data[x][1];
        }
      }
      
      if(!url) { continue; }
      
      var iter = MccApp.accounts()
      .withCondition('LabelNames CONTAINS "' + urlData[j][0] + '"')
      .get();
      
      var labels = [], names = [];
      while(iter.hasNext()) {
        var account = iter.next();
        MccApp.select(account);
        names.push([account.getName()]);
        
        var it = AdWordsApp.labels().get();
        while(it.hasNext()) {
          labels.push([account.getName(), it.next().getName()]);
        }
      }
      
      var ss = SpreadsheetApp.openByUrl(url);
      try {
        var labelTab = ss.getSheetByName('Labels Extract');
        labelTab.getRange(2,1,labelTab.getLastRow(),labelTab.getLastColumn()).clearContent();
        labelTab.getRange(labelTab.getLastRow()+1, 1, labels.length, labels[0].length).setValues(labels);
        labelTab.sort(1);
        
        var tab = ss.getSheetByName('Labels');
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(names).build();
        tab.getRange('A2:A').setDataValidation(rule);   
      } catch(ex) {
        Logger.log(ex);
      }
    }
    
    
  }