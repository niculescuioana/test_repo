function main() {
  
    var FROM = getAdWordsFormattedDate(90, 'yyyyMMdd');
    var TO = getAdWordsFormattedDate(1, 'yyyyMMdd');
    var URL = 'https://docs.google.com/spreadsheets/d/18g5xgo0VkwGM3ho1J1iElScrexem7uZLma7XFPOXwdU/edit#gid=0';
    var TAB_NAME = 'Verticals';
    
    var currencyMap = readCurrencyExchangeRates();
    currencyMap['GBP'] = 1;
    var tab = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
    var data = tab.getDataRange().getValues();
    data.shift();
    
    var map = {};
    for(var x in data) {
      map[data[x][1]] = parseInt(x,10);
    }
    
    var today = Utilities.formatDate(new Date(), 'GMT', 'MMM dd, yyyy');
    var year = parseInt(Utilities.formatDate(new Date(), 'GMT', 'yyyy'), 10);
    
    var TY_FROM = (year + '0101'),
        TY_TO = (year + '1231'),
        PY_FROM = ((year-1) + '0101'),
        PY_TO = ((year-1) + '1231'),
        PPY_FROM = ((year-2) + '0101'),
        PPY_TO = ((year-2) + '1231'),
        PPPY_FROM = ((year-3) + '0101'),
        PPPY_TO = ((year-3) + '1231');
    
    
    var temp = {};
    var MASTER_URL = 'https://docs.google.com/spreadsheets/d/16hkDsJ-K2LY0QzvcAGlsTTzGidW_Uk6C064wzY7OiF4/edit';
    var input = SpreadsheetApp.openByUrl(MASTER_URL).getSheetByName('Dashboard Urls').getDataRange().getValues();
    input.shift();
    
    for(var z in input) {
      if(input[z][0] == 'New Business') { continue; }
      
      var iter = MccApp.accounts()
      .withCondition('LabelNames CONTAINS "' + input[z][0] + '"')
      .withCondition('Impressions > 0')
      .forDateRange(FROM, TO)
      .get();
      
      while(iter.hasNext()) {
        var acc = iter.next();
        if(!acc.getName()) { continue; }
        MccApp.select(acc);
        
        var id = acc.getCustomerId();
        if(temp[id]) { continue; }
        
        temp[id] = 1;
        
        var currency = AdWordsApp.currentAccount().getCurrencyCode();
        
        var stats = AdWordsApp.currentAccount().getStatsFor(FROM, TO);
        var spend = stats.getCost()*currencyMap[currency];
        
        var cpa = stats.getConversions() == 0 ? 0 : spend / stats.getConversions();
        
        
        var index = map[id];
        if(!index) {
          data.push([
            acc.getName(), id, input[z][0], '', '', 
            
            (AdWordsApp.currentAccount().getStatsFor(PPPY_FROM, PPPY_TO).getCost()*currencyMap[currency]),
            (AdWordsApp.currentAccount().getStatsFor(PPY_FROM, PPY_TO).getCost()*currencyMap[currency]),
            (AdWordsApp.currentAccount().getStatsFor(PY_FROM, PY_TO).getCost()*currencyMap[currency]),
            (AdWordsApp.currentAccount().getStatsFor(TY_FROM, TY_TO).getCost()*currencyMap[currency]),
            
            stats.getClicks(), stats.getImpressions(),
            stats.getCtr(), stats.getAverageCpc(), currency,
            stats.getConversions(), cpa, stats.getConversionRate(), 
            stats.getCost(), spend, today
          ]);
          
          continue;
        }
        
        data[index][2] = input[z][0];
        data[index][5] = AdWordsApp.currentAccount().getStatsFor(PPPY_FROM, PPPY_TO).getCost();
        data[index][6] = AdWordsApp.currentAccount().getStatsFor(PPY_FROM, PPY_TO).getCost();
        data[index][7] = AdWordsApp.currentAccount().getStatsFor(PY_FROM, PY_TO).getCost();
        data[index][8] = AdWordsApp.currentAccount().getStatsFor(TY_FROM, TY_TO).getCost();
        
        data[index][9] = stats.getClicks();
        data[index][10] = stats.getImpressions();
        data[index][11] = stats.getCtr();
        data[index][12] = stats.getAverageCpc();
        data[index][13] = currency;
        data[index][14] = stats.getConversions();
        data[index][15] = cpa
        data[index][16] = stats.getConversionRate();
        data[index][17] = stats.getCost();
        data[index][18] = spend;
        
        data[index][19] = today;
      }
    }
    
    tab.getRange(2,1,data.length,data[0].length).setValues(data);
  }
  
  
  function getAdWordsFormattedDate(d, format){
    var date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(date,AdWordsApp.currentAccount().getTimeZone(),format);
  } 
  
  function readCurrencyExchangeRates() {
    var MCC_REPORT_URL = 'https://docs.google.com/spreadsheets/d/1F3bjn411jR3aYEJpLNAeCdqobJYlMQRUlK-6KVcdjCE/edit#gid=112312249';
    var CURRENCY_EXCHANGE_TAB_NAME = 'Currency Exchange';
    
    var map = {};
    var data = SpreadsheetApp.openByUrl(MCC_REPORT_URL).getSheetByName(CURRENCY_EXCHANGE_TAB_NAME).getDataRange().getValues();
    data.shift();
    
    for(var k in data) {
      if(!data[k][0]) { continue; }
      map[data[k][0]] = data[k][1]; 
    }
    
    return map;
  }
  