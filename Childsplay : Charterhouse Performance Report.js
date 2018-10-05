var CONFIG = {
    '309-121-1424': {
      'REPORT_URL': 'https://docs.google.com/spreadsheets/d/1mRlcU8YYyp9ARs2Sr3lAw1CwC1MqkvdIuYfNej8JQ_c/edit',
      'PROFILE': '15044218',
      'ROAS': 4
    },
    '954-734-3289': {
      'REPORT_URL': 'https://docs.google.com/spreadsheets/d/1udigqdz0fsuTHXVEfobgJEweuLjgEwdwjjzXG_-XPZI/edit',
      'PROFILE': '163505384',
      'ROAS': 3
    }
  }
  
  function main() {
    var ids = Object.keys(CONFIG);
    MccApp.accounts().withIds(ids).executeInParallel('compileReport');
  }
  
  function getLabelDetails(SETTINGS) {
    var labelMap = {}, labelKeyMap = {};
    var query = 'SELECT CampaignName,Labels FROM CAMPAIGN_PERFORMANCE_REPORT during TODAY';
    var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
    while(rows.hasNext()) {
     var row = rows.next();
      if(row.Labels == '--') { continue; }
      var labels = JSON.parse(row.Labels);
      
      labelKeyMap[row.CampaignName] = {}
      for(var z in labels) {
        if(!labelMap[labels[z]]) { labelMap[labels[z]] = {}; }
        labelMap[labels[z]][row.CampaignName]= 1;
        labelKeyMap[row.CampaignName][labels[z]] = 1;
      }
    }
    
    var query = 'SELECT CampaignName,AdGroupName,Labels FROM ADGROUP_PERFORMANCE_REPORT during TODAY';
    var rows = AdWordsApp.report(query, {'includeZeroImpressions': true}).rows();
    while(rows.hasNext()) {
     var row = rows.next();
      if(row.Labels == '--') { continue; }
      var labels = JSON.parse(row.Labels);
      
      var key = [row.CampaignName, row.AdGroupName].join('!~!');
      labelKeyMap[key] = {}
      for(var z in labels) {
        if(!labelMap[labels[z]]) { labelMap[labels[z]] = {}; }      
        labelMap[labels[z]][key]= 1;
        labelKeyMap[key][labels[z]] = 1;      
      }
    }
    
    
    for(var label in labelMap) {
      if(!Object.keys(labelMap[label]).length) {
        delete labelMap[label];
      }
    }
    
    addMissingLabelToSheet(labelMap, SETTINGS);
    
    return {'labelKeyMap': labelKeyMap, 'labelMap': labelMap };
  }
  
  function addMissingLabelToSheet(labelMap, SETTINGS) {
    var map = JSON.parse(JSON.stringify(labelMap));
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('Labels');
    var data = tab.getDataRange().getValues();
    data.shift();
    
    for(var z in data) {
      if(data[z][1] != 'Yes') {
        delete labelMap[data[z][0]];
      }
      
      delete map[data[z][0]];
    }
    
    var out = [];
    for(var label in map) {
      out.push([label]);
    }
    
    if(out.length) {
      tab.getRange(tab.getLastRow()+1,1,out.length,1).setValues(out);
    }
  }
  
  function compileReport() {
    var SETTINGS = CONFIG[AdWordsApp.currentAccount().getCustomerId()];
    var labelDataMap = getLabelDetails(SETTINGS);
  
    var optArgs = { 'dimensions': 'ga:campaign,ga:adGroup', 'filters': 'ga:medium==cpc;ga:source==google', 'max-results': '10000' };
    
    // Month to Date DATA Begins
    
    var date = new Date(getAdWordsFormattedDate_(1, 'MMM d, yyyy'));
    var diff = date.getDate();
    
    var stats = {}, agStats = {};
    getDataFromAnalytics_(stats, agStats, getAdWordsFormattedDate_(diff, 'yyyy-MM-dd'), getAdWordsFormattedDate_(1, 'yyyy-MM-dd'), optArgs, labelDataMap, SETTINGS.PROFILE);
    getDataFromMCF_(stats, agStats, getAdWordsFormattedDate_(diff, 'yyyy-MM-dd'), getAdWordsFormattedDate_(1, 'yyyy-MM-dd'), labelDataMap, SETTINGS.PROFILE);
    
    //getDataFromAnalytics_(stats, agStats, '2017-11-01', '2017-11-30', optArgs, labelDataMap);
    //getDataFromMCF_(stats, agStats, '2017-11-01', '2017-11-30', labelDataMap);
    
    var out = [], categoryStats = {
      'Branding': { 'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0 },
      'Non Branding': { 'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0 }
    };
    for(var key in stats) {
      var row = stats[key];
      
      var category = 'Branding';
      if(key != 'Branding') {
        category = 'Non Branding';
      } 
    
      for(var metric in row) {
        categoryStats[category][metric] += row[metric]; 
      }
      
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4);
      row.AssistedROAS = row.Cost == 0 ? 0 : round_((row.ConversionValue+row.AssistedRevenue) / row.Cost, 4);
      row.CPC = row.Clicks == 0 ? 0 : round_(row.Cost / row.Clicks, 4);
      
      row.CPR = row.ConversionValue == 0 ? 0 : round_(row.Cost / row.ConversionValue, 4);
      row.AssistedCPR = (row.ConversionValue+row.AssistedRevenue) == 0 ? 0 : round_(row.Cost / (row.ConversionValue+row.AssistedRevenue), 4);
      
      out.push([
        key, row.Impressions, row.Clicks, row.Conversions, row.Cost, row.CPC, row.ConversionValue, row.AssistedRevenue, 
        row.ROAS, row.AssistedROAS, row.CPR, row.AssistedCPR
      ]);
    }
    
    
    out.sort(function(a,b) {return b[4] - a[4];});
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('Summary');
    tab.getRange(11, 1, tab.getLastRow()-9, tab.getLastColumn()).clearContent();
    if(out.length) {
      tab.getRange(11, 1, out.length, out[0].length).setValues(out);
    }
    
    var out = [];
    for(var category in categoryStats) {
      var row = categoryStats[category];
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4);
      row.AssistedROAS = row.Cost == 0 ? 0 : round_((row.ConversionValue+row.AssistedRevenue) / row.Cost, 4);
      row.CPC = row.Clicks == 0 ? 0 : round_(row.Cost / row.Clicks, 4);
      
      row.CPR = row.ConversionValue == 0 ? 0 : round_(row.Cost / row.ConversionValue, 4);
      row.AssistedCPR = (row.ConversionValue+row.AssistedRevenue) == 0 ? 0 : round_(row.Cost / (row.ConversionValue+row.AssistedRevenue), 4);
      
      out.push([
        category, row.Impressions, row.Clicks, row.Conversions, row.Cost, row.CPC, row.ConversionValue, row.AssistedRevenue, 
        row.ROAS, row.AssistedROAS, row.CPR, row.AssistedCPR
      ]);
    }
    
    tab.getRange(6, 1, out.length, out[0].length).setValues(out);
    
    
    //return;
    // LAST 7 DAYS DATA Begins
    
    
    var stats = {}, agStats = {};
    getDataFromAnalytics_(stats, agStats, getAdWordsFormattedDate_(7, 'yyyy-MM-dd'), getAdWordsFormattedDate_(1, 'yyyy-MM-dd'), optArgs, labelDataMap, SETTINGS.PROFILE);
    getDataFromMCF_(stats, agStats, getAdWordsFormattedDate_(7, 'yyyy-MM-dd'), getAdWordsFormattedDate_(1, 'yyyy-MM-dd'), labelDataMap, SETTINGS.PROFILE);
    
    var out = [], categoryStats = {
      'Branding': { 'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0 },
      'Non Branding': { 'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0 }
    };
    for(var key in stats) {
      var row = stats[key];
      
      var category = 'Branding';
      if(key != 'Branding') {
        category = 'Non Branding';
      } 
    
      for(var metric in row) {
        categoryStats[category][metric] += row[metric]; 
      }
      
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4)
      row.AssistedROAS = row.Cost == 0 ? 0 : round_((row.ConversionValue+row.AssistedRevenue) / row.Cost, 4)
      
      out.push([key, row.Impressions, row.Clicks, row.Conversions, row.Cost, row.ConversionValue, row.AssistedRevenue, row.ROAS, row.AssistedROAS]);
    }
    
    
    out.sort(function(a,b) {return b[4] - a[4];});
    
    var out_2 = [], out_3 = [];
    for(var z in out) {
      if(out[z][7] < SETTINGS.ROAS && out_2.length < 10) {
        out_2.push(out[z]);
      } else if(out[z][7] >= SETTINGS.ROAS && out_3.length < 10) {
        out_3.push(out[z]);
      }
    }
    
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('Brand Summary');
    tab.getRange(3, 1, 10, tab.getLastColumn()).clearContent();
    if(out_3.length) {
      tab.getRange(3, 1, out_3.length, out_3[0].length).setValues(out_3);
    }
    
    tab.getRange(16, 1, 10, tab.getLastColumn()).clearContent();  
    if(out_2.length) {
      tab.getRange(16, 1, out_2.length, out_2[0].length).setValues(out_2);
    }
    
    
    var out = [];
    for(var category in categoryStats) {
      var row = categoryStats[category];
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4);
      row.AssistedROAS = row.Cost == 0 ? 0 : round_((row.ConversionValue+row.AssistedRevenue) / row.Cost, 4);
      row.CPC = row.Clicks == 0 ? 0 : round_(row.Cost / row.Clicks, 4);
      
      row.CPR = row.ConversionValue == 0 ? 0 : round_(row.Cost / row.ConversionValue, 4);
      row.AssistedCPR = (row.ConversionValue+row.AssistedRevenue) == 0 ? 0 : round_(row.Cost / (row.ConversionValue+row.AssistedRevenue), 4);
      
      out.push([
        category, row.Impressions, row.Clicks, row.Conversions, row.Cost, row.CPC, row.ConversionValue, row.AssistedRevenue, 
        row.ROAS, row.AssistedROAS, row.CPR, row.AssistedCPR
      ]);
    }
    
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('Summary');
    tab.getRange(2, 1, out.length, out[0].length).setValues(out);
    
    
    var out = [];
    for(var key in agStats) {
      var row = agStats[key];
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4);
      row.AssistedROAS = row.Cost == 0 ? 0 : round_((row.ConversionValue+row.AssistedRevenue) / row.Cost, 4);
      out.push(key.split('!~!').concat([row.Impressions, row.Clicks, row.Conversions, row.Cost, row.ConversionValue, row.AssistedRevenue, row.ROAS, row.AssistedROAS]));
    }
    
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('AdGroup Summary');
    tab.getRange(3,1,tab.getLastRow(),tab.getLastColumn()).clearContent();
    tab.getRange(3,1,out.length,out[0].length).setValues(out);
    tab.sort(5, false);  
    
    var statsOld = {}, agStatsOld = {};
    getDataFromAnalytics_(statsOld, agStatsOld, getAdWordsFormattedDate_(14, 'yyyy-MM-dd'), getAdWordsFormattedDate_(8, 'yyyy-MM-dd'), optArgs, labelDataMap, SETTINGS.PROFILE);
    
    for(var key in statsOld) {
      var row = statsOld[key];
      row.ROAS = row.Cost == 0 ? 0 : round_(row.ConversionValue / row.Cost, 4);
    }
    
    var out = [];
    for(var key in stats) {
      var row = stats[key];
      var lastRow = statsOld[key];
      if(!lastRow) {
        lastRow = {
          'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0, 'ROAS': 0
        }; 
      }
      
      var change  = getChange_(row.ROAS, lastRow.ROAS);
      out.push([key, row.Clicks, lastRow.Clicks, row.Cost, lastRow.Cost, row.ConversionValue, lastRow.Clicks, row.ROAS, lastRow.ROAS, change]);
    }
    
    out.sort(function(a,b) {return b[9] - a[9];});
    
    var tab = SpreadsheetApp.openByUrl(SETTINGS.REPORT_URL).getSheetByName('Brand Summary');
    var out_2 = [];
    for(var z in out) {
      if(out_2.length == 5) { break; }
      if(out[z][9] <= 0) { break; }
      out_2.push(out[z]);
    }
    
    tab.getRange(30,1,5,tab.getLastColumn()).clearContent();
    if(out_2.length) {
      tab.getRange(30,1,out_2.length,out_2[0].length).setValues(out_2);
    }
    
    
    var out_3 = [];
    for(var z=out.length-1; z>=0; z--) {
      if(out_3.length == 5) { break; }
      if(out[z][9] >= 0) { break; }
      out_3.push(out[z]);
    }
    
    tab.getRange(39,1,5,tab.getLastColumn()).clearContent();
    if(out_3.length) {
      tab.getRange(39,1,out_3.length,out_3[0].length).setValues(out_3);
    }
  }
  
  
  function getDataFromAnalytics_(stats,agStats,FROM,TO,optArgs,labelDataMap,PROFILE_ID) {
    var labelKeyMap = labelDataMap['labelKeyMap'],
        labelMap = labelDataMap['labelMap'];
    
    var attempts = 3;
    
    // Make a request to the API.
    while(attempts > 0) {
      try {
        var resp = Analytics.Data.Ga.get(
          'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
          FROM,                 // Start-date (format yyyy-MM-dd).
          TO,                  // End-date (format yyyy-MM-dd).
          "ga:adCost,ga:adClicks,ga:transactions,ga:transactionRevenue,ga:impressions",
          optArgs);
        
        break;
      } catch(ex) {
        Logger.log(ex);
        attempts--;
        Utilities.sleep(2500);
      }
    }
    
    var initMap = {
      'Cost': 0, 'Clicks': 0, 'Conversions': 0, 'ConversionValue': 0, 'Impressions': 0, 'AssistedRevenue': 0
    };
    
    var rows = resp.getRows();
    for(var k in rows) {
      var camp = rows[k][0];
      if(labelKeyMap[camp]) {
        for(var key in labelKeyMap[camp]) {
          if(!labelMap[key]) { continue; }
          
          if(!stats[key]) {
            stats[key] = JSON.parse(JSON.stringify(initMap));
          }
          
          stats[key].Cost += parseFloat(rows[k][2]);
          stats[key].Clicks += parseInt(rows[k][3],10);    
          stats[key].Conversions += parseInt(rows[k][4],10);
          stats[key].ConversionValue += parseFloat(rows[k][5]);
          stats[key].Impressions += parseInt(rows[k][6],10);           
        }
      }
      
      var camp = [rows[k][0],rows[k][1]].join('!~!')
      if(labelKeyMap[camp]) {
        for(var key in labelKeyMap[camp]) {
          if(!labelMap[key]) { continue; }
          
          if(!stats[key]) {
            stats[key] = JSON.parse(JSON.stringify(initMap));
          }
          
          stats[key].Cost += parseFloat(rows[k][2]);
          stats[key].Clicks += parseInt(rows[k][3],10);    
          stats[key].Conversions += parseInt(rows[k][4],10);
          stats[key].ConversionValue += parseFloat(rows[k][5]);
          stats[key].Impressions += parseInt(rows[k][6],10);           
        }
      }
      
      var agKey = [rows[k][0],rows[k][1]].join('!~!')
      if(!agStats[agKey]) {
        agStats[agKey] = JSON.parse(JSON.stringify(initMap));
      }
      
      agStats[agKey].Cost += parseFloat(rows[k][2]);
      agStats[agKey].Clicks += parseInt(rows[k][3],10);    
      agStats[agKey].Conversions += parseInt(rows[k][4],10);
      agStats[agKey].ConversionValue += parseFloat(rows[k][5]);
      agStats[agKey].Impressions += parseInt(rows[k][6],10);     
    }
  }
  
  function getDataFromMCF_(stats,agStats,FROM,TO,labelDataMap,PROFILE_ID) {
    var labelKeyMap = labelDataMap['labelKeyMap'],
        labelMap = labelDataMap['labelMap'];
    var filters = ['mcf:basicChannelGroupingPath=@Paid Search','mcf:basicChannelGroupingPath!=Paid Search',
                   'mcf:conversionType==Transaction'];
    var optArgs = { 
      'dimensions': 'mcf:basicChannelGroupingPath,mcf:adwordsCampaignPath,mcf:adwordsAdGroupPath',
      'filters': filters.join(';'), 'max-results': '10000'
    };
    
    try {
      var results = Analytics.Data.Mcf.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "mcf:totalConversionValue,mcf:totalConversions",
        optArgs
      );
    } catch(ex) {
      Logger.log(ex);
      return;
    }
    
    var rows = results.rows;
    for(var k in rows) {
      var channelGroups = JSON.parse(rows[k][0]).conversionPathValue;
      if(channelGroups[0].nodeValue != 'Paid Search') { continue; }
      var index = channelGroups.length-1;
      if(channelGroups[index].nodeValue == 'Paid Search') { continue; }
      
      var camp = JSON.parse(rows[k][1]).conversionPathValue[0].nodeValue;
      if(camp == '(unavailable)') {
        camp = '(not set)';
      }
      
      var ag = JSON.parse(rows[k][2]).conversionPathValue[0].nodeValue;
      if(ag == '(unavailable)') {
        ag = '(not set)';
      }
      
      if(labelKeyMap[camp]) {
        for(var key in labelKeyMap[camp]) {
          if(!labelMap[key] || !stats[key]) { continue; }
          stats[key].AssistedRevenue += parseFloat(rows[k][3].primitiveValue);
        }
      }
      
      var agKey = [camp,ag].join('!~!');
      if(labelKeyMap[agKey]) {
        for(var key in labelKeyMap[agKey]) {
          if(!labelMap[key] || !stats[key]) { continue; }
          stats[key].AssistedRevenue += parseFloat(rows[k][3].primitiveValue);
        }
      }
      
      var key = [camp,ag].join('!~!');
      if(!agStats[key]) { continue; }
      agStats[key].AssistedRevenue += parseFloat(rows[k][3].primitiveValue);
      
    }
  }
  
  function getAdWordsFormattedDate_(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,'GMT',format);
  }
  
  
  function round_(num,n) {    
    return +(Math.round(num + "e+"+n)  + "e-"+n);
  }
  
  
  function getChange_(a,b) {
    if(b == 0 && a == 0) { return 0;}
    if(b == 0 && a > 0) { return 1; }
    
    return (a-b)/b;
  }