var PROFILE_ID_MAP = {
    '119026770': {
      'EMAIL': 'awais@tackville.co.uk,steff@tackville.co.uk,neeraj@pushgroup.co.uk,ian@pushgroup.co.uk',
      //'EMAIL': 'naman@pushgroup.co.uk,neeraj@pushgroup.co.uk',
      'ACCOUNT_NAME': 'Tackville'
    }
  };
  
  function main() {
    var day = Utilities.formatDate(new Date(), 'GMT', 'EEE');
    
    for(var id in PROFILE_ID_MAP) {
      if(day != 'Fri' && PROFILE_ID_MAP[id].ACCOUNT_NAME == 'Tackville') {
        continue; 
      }
      
      runScript(id, PROFILE_ID_MAP[id].EMAIL, PROFILE_ID_MAP[id].ACCOUNT_NAME); 
    }
  }
  
  function runScript(PROFILE_ID, EMAIL, ACCOUNT_NAME) {
    var today = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
    today.setHours(12);
    today.setDate(1);
    var FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
    var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
    
    var stats = { 
      'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'YTCost': 0
    };
    
    getDataFromAnalytics(PROFILE_ID,stats,FROM,TO);
    getDataFromMCF(PROFILE_ID,stats,FROM,TO);
    
    stats.CPA = stats.Conversions == 0 ? 0 : round(stats.Cost/stats.Conversions,2);
    stats.ROAS = stats.Cost == 0 ? '0.00%' : round(100*stats.ConversionValue/stats.Cost,2)+'%';
    
    var MSG = 'Hi '+ACCOUNT_NAME+',\n\n';
    MSG += 'Please see the updated report below for month to date:\n\n'
    
    if(ACCOUNT_NAME == 'Letteroom') { 
      MSG += 'PPC Spend (Excluding Youtube): £' + round(stats.Cost,2) + '\n';
    } else {
      MSG += 'PPC Spend: £' + round(stats.Cost,2) + '\n';
    }
    
    MSG += 'Conversions: ' + stats.Conversions + '\n';
    MSG += 'Revenue: £' + round(stats.ConversionValue,2) + '\n';
    MSG += 'ROAS: ' + stats.ROAS + '\n';  
    MSG += 'CPA: ' + stats.CPA + '\n\n';  
    
    if(stats.YTCost > 0) {
      MSG += 'Youtube Spend: £' + round(stats.YTCost,2) + '\n\n';  
    }
    
    MSG += 'Kind Regards,\nPush';
    
    MailApp.sendEmail(EMAIL, ACCOUNT_NAME + ' - MTD Report',MSG);
  }
  
  function getDataFromAnalytics(PROFILE_ID,stats,FROM,TO) {
    var optArgs = { 
      'dimensions': 'ga:campaign', 
      'filters': 'ga:medium==cpc;ga:source==google' 
    };
    
    var attempts = 3;
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
      if(rows[k][0] == 'GV - TV Shopping') { 
        stats.YTCost += parseFloat(rows[k][1]);
      } else {
        stats.Cost += parseFloat(rows[k][1]);
      }
      stats.Conversions += parseInt(rows[k][2],10);
      stats.ConversionValue += parseFloat(rows[k][3]);
    }
  }
  
  function getDataFromMCF(PROFILE_ID,stats,FROM, TO) {
    var filters = ['mcf:basicChannelGroupingPath=@Paid Search',
                     'mcf:basicChannelGroupingPath!=Paid Search',
                   'mcf:conversionType==Transaction'];
    
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
  
  function round(num,n) {    
    return +(Math.round(num + "e+"+n)  + "e-"+n);
  }