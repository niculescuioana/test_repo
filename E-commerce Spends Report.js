var LABEL = 'Ian';
function main() {
  
  var LABELS = [
    'Neeraj','Ian','Mike','Jay'
  ];

  
  var CURRENCY_MAP = readCurrencyExchangeRates();
  CURRENCY_MAP['GBP'] = 1;
  
  MccApp.accounts().withCondition('LabelNames CONTAINS "'+LABEL+'"').executeInParallel('run','compile',JSON.stringify(CURRENCY_MAP));
    
}

function compile(results) {
  var URL = 'https://docs.google.com/spreadsheets/d/1TqfaP-MPLlxsQpaNK1G_zw4aQC42QG_SvEY_peIuhHY/edit#gid=809402620';
  var TAB_NAME = 'Q2 2018';
  
  var tab = SpreadsheetApp.openByUrl(URL).getSheetByName(TAB_NAME);
  
  var out = [];
  for(var i in results) {
   out.push(JSON.parse(results[i].getReturnValue())); 
  }
  
  tab.getRange(tab.getLastRow()+1,1,out.length,out[0].length).setValues(out);
}

function run(input) {
  var CURRENCY_MAP = JSON.parse(input);
  
  var start = '20180401';
  var end = '20180630';
  
  
      var query = [
        'SELECT CampaignId, AdvertisingChannelType, Cost, AccountCurrencyCode FROM CAMPAIGN_PERFORMANCE_REPORT',
        'WHERE Cost > 0',
        'DURING', start + ',' + end
      ].join(' ');
      
      var shoppingSpends = 0, nonshoppingSpends = 0;
      var rows = AdWordsApp.report(query).rows();
      while(rows.hasNext()) {
        var row = rows.next();
        row.Cost = CURRENCY_MAP[row.AccountCurrencyCode] * parseFloat(row.Cost.toString().replace(/,/g, ''));
        
        if(row.AdvertisingChannelType == 'Shopping') {
          shoppingSpends += row.Cost;
        } else {
          nonshoppingSpends += row.Cost; 
        }
      }
    
  return JSON.stringify([AdWordsApp.currentAccount().getCustomerId(), AdWordsApp.currentAccount().getName(), LABEL, shoppingSpends, nonshoppingSpends]);
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
