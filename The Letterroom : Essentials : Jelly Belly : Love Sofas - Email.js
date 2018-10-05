function main() {
  var date = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
  var day = date.getDay();
  try {
    sendReportToJellyBelly();
  } catch(ex) {
    Logger.log('sendReportToJellyBelly: ' + ex);    
  }
  
  try {
    sendReportToLetterRoom();
  } catch(ex) {
    Logger.log('sendReportToLetterRoom: ' + ex);    
  }
  
  try {
    //sendReportToEssentials();
  } catch(ex) {
    Logger.log('sendReportToEssentials: ' + ex); 
  }
  
  if(day == 1) {
    try {
      sendReportToLoveSofas();
    } catch(ex) {
      Logger.log('sendReportToLoveSofas: ' + ex); 
    }
  }
}

function sendReportToLoveSofas() {
  MccApp.select(MccApp.accounts().withIds(['904-921-2976']).get().next());  
  var ACCOUNT_NAME = 'Love Sofas';
  var PROFILE_ID = 67578528;
  
  var today = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
  today.setHours(12);
  today.setDate(1);
  var FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var stats = {};
  
  var optArgs = { 
    'filters': 'ga:medium==cpc;ga:source==google' 
  };
  
  var attempts = 3;
  while(attempts > 0) {
    try {
      var resp = Analytics.Data.Ga.get(
        'ga:'+PROFILE_ID,      // Table id (format ga:xxxxxx).
        FROM,                 // Start-date (format yyyy-MM-dd).
        TO,                  // End-date (format yyyy-MM-dd).
        "ga:adCost,ga:transactions,ga:transactionRevenue,ga:adClicks,ga:goal2Completions,ga:goal2Value",
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
    stats = { 
      'Cost': parseFloat(rows[k][0]),
      'Conversions': parseInt(rows[k][1],10),
      'ConversionValue': parseFloat(rows[k][2]),
      'Clicks': parseInt(rows[k][3],10),
      'Goals': parseInt(rows[k][4],10),
      'GoalValue': parseFloat(rows[k][5])
    }
  }
  
  var mcfStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  getDataFromMCF(PROFILE_ID,mcfStats,FROM,TO);
  
  var totalRevenue = (stats['ConversionValue']+mcfStats['ConversionValue']+stats['GoalValue']);
  var ROAS = round(100*totalRevenue / stats['Cost'], 2);
    
  var MSG = 'Hi Love Sofas,\n\n';
  MSG += 'Please see the updated report below for PPC month to date:\n\n'
  
  MSG += 'PPC Spend: £' + stats['Cost'] + '\n\n';
  
  MSG += 'Conversions: ' + (stats['Conversions']+mcfStats['Conversions']) + '\n';
  MSG += 'Revenue: £' + round((stats['ConversionValue']+mcfStats['ConversionValue']),2) + '\n\n';
  
  MSG += 'Number Of Phone Calls: ' + stats['Goals'] + '\n';  
  MSG += 'Revenue from Phone Calls: £' + stats['GoalValue'] + '\n\n';  
  
  MSG += 'Total Revenue: £' + round(totalRevenue,2) + '\n\n';
  MSG += 'ROAS: ' + ROAS + '%\n\n';  
  MSG += 'Kind Regards,\nPush';
  

  var EMAIL = 'paulina@lovesofas.co.uk,richard@lovesofas.co.uk,neeraj@pushgroup.co.uk,mike@pushgroup.co.uk';
  //var EMAIL = 'naman@pushgroup.co.uk';
  MailApp.sendEmail(EMAIL, ACCOUNT_NAME + ' - MTD Report', MSG);
}

function sendReportToJellyBelly() {
  MccApp.select(MccApp.accounts().withIds(['284-674-7559']).get().next());  
  var ACCOUNT_NAME = 'Jelly Belly';
  var PROFILE_ID = 70632539;
  
  var today = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
  today.setHours(12);
  today.setDate(1);
  var FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var campaignStats = getDataFromAnalytics(PROFILE_ID,FROM,TO);
  
  var stats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  getDataFromMCF(PROFILE_ID,stats,FROM,TO);
  
  
  var searchStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var brandStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var iter = AdWordsApp.campaigns()
  .withCondition('AdvertisingChannelType = SEARCH')
  .get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      if(campName == "Branding") {
        brandStats.Cost += campStats.Cost;
        brandStats.Conversions += campStats.Conversions;
        brandStats.ConversionValue += campStats.ConversionValue;
      } else {
        searchStats.Cost += campStats.Cost;
        searchStats.Conversions += campStats.Conversions;
        searchStats.ConversionValue += campStats.ConversionValue;
      }
    }
  }
  var shoppingStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var iter = AdWordsApp.shoppingCampaigns().get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      shoppingStats.Cost += campStats.Cost;
      shoppingStats.Conversions += campStats.Conversions;
      shoppingStats.ConversionValue += campStats.ConversionValue;
    }
  }
  
  shoppingStats.Conversions += stats.Conversions;
  shoppingStats.ConversionValue += stats.ConversionValue; 
  
  
  brandStats.CPA = brandStats.Conversions == 0 ? 0 : round(brandStats.Cost/brandStats.Conversions,2);
  brandStats.ROAS = brandStats.Cost == 0 ? '0.00%' : round(100*brandStats.ConversionValue/brandStats.Cost,2)+'%';
  
  searchStats.CPA = searchStats.Conversions == 0 ? 0 : round(searchStats.Cost/searchStats.Conversions,2);
  searchStats.ROAS = searchStats.Cost == 0 ? '0.00%' : round(100*searchStats.ConversionValue/searchStats.Cost,2)+'%';
  
  shoppingStats.CPA = shoppingStats.Conversions == 0 ? 0 : round(shoppingStats.Cost/shoppingStats.Conversions,2);
  shoppingStats.ROAS = shoppingStats.Cost == 0 ? '0.00%' : round(100*shoppingStats.ConversionValue/shoppingStats.Cost,2)+'%';
  
  
  var MSG = 'Hi Jelly Belly,\n\n';
  MSG += 'Please see the updated report below for month to date:\n\n'
  
  MSG += 'Shopping:\n';
  MSG += 'PPC Spend: £' + round(shoppingStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + shoppingStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(shoppingStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + shoppingStats.ROAS + '\n';  
  MSG += 'CPA: £' + shoppingStats.CPA + '\n\n';  
  
  MSG += 'Search (Branding):\n';
  MSG += 'PPC Spend: £' + round(brandStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + brandStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(brandStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + brandStats.ROAS + '\n';  
  MSG += 'CPA: £' + brandStats.CPA + '\n\n';  
  
  MSG += 'Search (Other):\n';
  MSG += 'PPC Spend: £' + round(searchStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + searchStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(searchStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + searchStats.ROAS + '\n';  
  MSG += 'CPA: £' + searchStats.CPA + '\n\n';  
  
  MSG += 'Kind Regards,\nPush';
  
  
  var EMAIL = 'talfano@bestimports.co.uk,anniem@bestimports.co.uk,victoria@jellybelly.co.uk,victoria@jellybelly-uk.com,mike@pushgroup.co.uk';
  //EMAIL = 'naman@pushgroup.co.uk';
  MailApp.sendEmail(EMAIL, ACCOUNT_NAME + ' - MTD Report', MSG);
}

function sendReportToEssentials() {
  MccApp.select(MccApp.accounts().withIds(['564-205-6978']).get().next());
  var ACCOUNT_NAME = 'Essentials London';
  var PROFILE_ID = 14426663;
  
  var today = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
  today.setHours(12);
  today.setDate(1);
  var FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var campaignStats = getDataFromAnalytics(PROFILE_ID,FROM,TO);
  
  var searchStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0, 'SIS': '-'
  };
  
  var iter = AdWordsApp.campaigns()
  .withCondition('AdvertisingChannelType = SEARCH')
  .get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      searchStats.Cost += campStats.Cost;
      searchStats.Conversions += campStats.Conversions;
      searchStats.ConversionValue += campStats.ConversionValue;
      searchStats.Clicks += campStats.Clicks;
    }
  }
  var shoppingStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var iter = AdWordsApp.shoppingCampaigns().get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      shoppingStats.Cost += campStats.Cost;
      shoppingStats.Conversions += campStats.Conversions;
      shoppingStats.ConversionValue += campStats.ConversionValue;
      shoppingStats.Clicks += campStats.Clicks;      
    }
  }
  
  
  searchStats.CPC = searchStats.Clicks == 0 ? 0 : round(searchStats.Cost/searchStats.Clicks,2);
  searchStats.ROAS = searchStats.ConversionValue == 0 ? '0.00%' : round(100*searchStats.Cost/searchStats.ConversionValue,2)+'%';
  
  shoppingStats.CPC = shoppingStats.Clicks == 0 ? 0 : round(shoppingStats.Cost/shoppingStats.Clicks,2);
  shoppingStats.ROAS = shoppingStats.ConversionValue == 0 ? '0.00%' : round(100*shoppingStats.Cost/shoppingStats.ConversionValue,2)+'%';
  
  
  var query = 'SELECT SearchImpressionShare FROM ACCOUNT_PERFORMANCE_REPORT WHERE AdNetworkType1 = SEARCH DURING THIS_MONTH';
  var rows = AdWordsApp.report(query).rows();
  if(rows.hasNext()) {
    searchStats.SIS = rows.next()['SearchImpressionShare'];
  }
  
  var MSG = 'Hi Essentials,\n\n';
  MSG += 'Please see the updated report below for month to date:\n\n'
  
  MSG += 'Shopping:\n';
  MSG += 'PPC Spend: £' + round(shoppingStats.Cost,2) + '\n';
  MSG += 'Revenue: £' + round(shoppingStats.ConversionValue,2) + '\n';
  MSG += '% Of Cost vs Sales: ' + shoppingStats.ROAS + '\n';  
  MSG += 'CPC: £' + shoppingStats.CPC + '\n\n';  
  
  MSG += 'Search:\n';
  MSG += 'PPC Spend: £' + round(searchStats.Cost,2) + '\n';
  MSG += 'Revenue: £' + round(searchStats.ConversionValue,2) + '\n';
  MSG += 'CPC: £' + searchStats.CPC + '\n';  
  MSG += 'Search Impr Share: ' + searchStats.SIS + '\n\n';  
  
  MSG += 'Kind Regards,\nPush';
  
  var EMAIL = 'f.mobara@gmail.com,varavipour@gmail.com,neeraj@pushgroup.co.uk,jay@pushgroup.co.uk';
  //EMAIL = 'naman@pushgroup.co.uk,jay@pushgroup.co.uk';
  MailApp.sendEmail(EMAIL, ACCOUNT_NAME + ' - MTD Report', MSG);
}


function sendReportToLetterRoom() {
  MccApp.select(MccApp.accounts().withIds(['889-498-4604']).get().next());  
  var ACCOUNT_NAME = 'Letteroom';
  var PROFILE_ID = 38769215;
  
  var today = new Date(Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy'));
  today.setHours(12);
  today.setDate(1);
  var FROM = Utilities.formatDate(today, 'PST', 'yyyy-MM-dd');
  var TO = getAdWordsFormattedDate(0, 'yyyy-MM-dd');
  
  var campaignStats = getDataFromAnalytics(PROFILE_ID,FROM,TO);
  
  var stats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  getDataFromMCF(PROFILE_ID,stats,FROM,TO);
  
  
  var searchStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var brandStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var iter = AdWordsApp.campaigns()
  .withCondition('AdvertisingChannelType = SEARCH')
  .get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      if(campName == "Push - Branding") {
        brandStats.Cost += campStats.Cost;
        brandStats.Conversions += campStats.Conversions;
        brandStats.ConversionValue += campStats.ConversionValue;
      } else {
        searchStats.Cost += campStats.Cost;
        searchStats.Conversions += campStats.Conversions;
        searchStats.ConversionValue += campStats.ConversionValue;
      }
    }
  }
  var shoppingStats = { 
    'Cost': 0, 'Conversions': 0, 'ConversionValue': 0, 'Clicks': 0
  };
  
  var iter = AdWordsApp.shoppingCampaigns().get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      shoppingStats.Cost += campStats.Cost;
      shoppingStats.Conversions += campStats.Conversions;
      shoppingStats.ConversionValue += campStats.ConversionValue;
    }
  }
  
  shoppingStats.Conversions += stats.Conversions;
  shoppingStats.ConversionValue += stats.ConversionValue; 
  
  var ytStats = { 
    'Cost': 0, 'Views': 0
  };
  
  var iter = AdWordsApp.videoCampaigns().get();
  while(iter.hasNext()){
    var camp = iter.next(); 
    var campName = camp.getName();
    var campStats = campaignStats[campName];
    if(campStats) {
      ytStats.Cost += campStats.Cost;
    }
    
    ytStats.Views += camp.getStatsFor(FROM.replace(/-/g, ''), TO.replace(/-/g, '')).getViews();
  }
  
  brandStats.CPA = brandStats.Conversions == 0 ? 0 : round(brandStats.Cost/brandStats.Conversions,2);
  brandStats.ROAS = brandStats.Cost == 0 ? '0.00%' : round(100*brandStats.ConversionValue/brandStats.Cost,2)+'%';
  
  searchStats.CPA = searchStats.Conversions == 0 ? 0 : round(searchStats.Cost/searchStats.Conversions,2);
  searchStats.ROAS = searchStats.Cost == 0 ? '0.00%' : round(100*searchStats.ConversionValue/searchStats.Cost,2)+'%';
  
  shoppingStats.CPA = shoppingStats.Conversions == 0 ? 0 : round(shoppingStats.Cost/shoppingStats.Conversions,2);
  shoppingStats.ROAS = shoppingStats.Cost == 0 ? '0.00%' : round(100*shoppingStats.ConversionValue/shoppingStats.Cost,2)+'%';
  
  ytStats.CPV = ytStats.Views == 0 ? 0 : round(ytStats.Cost/ytStats.Views,2);
  
  var MSG = 'Hi Letteroom,\n\n';
  MSG += 'Please see the updated report below for month to date:\n\n'
  
  MSG += 'Shopping:\n';
  MSG += 'PPC Spend: £' + round(shoppingStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + shoppingStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(shoppingStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + shoppingStats.ROAS + '\n';  
  MSG += 'CPA: ' + shoppingStats.CPA + '\n\n';  
  
  MSG += 'Search (Branding):\n';
  MSG += 'PPC Spend: £' + round(brandStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + brandStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(brandStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + brandStats.ROAS + '\n';  
  MSG += 'CPA: ' + brandStats.CPA + '\n\n';  
  
  MSG += 'Search (Other):\n';
  MSG += 'PPC Spend: £' + round(searchStats.Cost,2) + '\n';
  MSG += 'Conversions: ' + searchStats.Conversions + '\n';
  MSG += 'Revenue: £' + round(searchStats.ConversionValue,2) + '\n';
  MSG += 'ROAS: ' + searchStats.ROAS + '\n';  
  MSG += 'CPA: ' + searchStats.CPA + '\n\n';  
  
  MSG += 'Youtube:\n';
  MSG += 'PPC Spend: £' + round(ytStats.Cost,2) + '\n';
  MSG += 'Views: ' + ytStats.Views + '\n';
  MSG += 'CPV (Cost Per View): ' + ytStats.CPV + '\n\n';
  
  MSG += 'Kind Regards,\nPush';
  
  var EMAIL = 'sherrie@theletteroom.com,jackie@theletteroom.com,kirsty@theletteroom.com,jay@pushgroup.co.uk';
  //EMAIL = 'naman@pushgroup.co.uk';
  MailApp.sendEmail(EMAIL, ACCOUNT_NAME + ' - MTD Report', MSG);
}

function getDataFromAnalytics(PROFILE_ID,FROM,TO) {
  var stats = {};
               
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
        "ga:adCost,ga:transactions,ga:transactionRevenue,ga:adClicks",
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
    stats[rows[k][0]] = { 
      'Cost': parseFloat(rows[k][1]),
      'Conversions': parseInt(rows[k][2],10),
      'ConversionValue': parseFloat(rows[k][3]),
      'Clicks': parseInt(rows[k][4],10)
    }
  }
  
  return stats;
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