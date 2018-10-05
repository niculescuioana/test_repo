var KEYWORD_CHECKER_URL = 'https://docs.google.com/spreadsheets/d/18yfACc_vAfDFtsHW9aI_n7iJaHAQt50oTAU1O_5AIM8/edit';

function main() {
  MccApp.accounts()
  .withCondition('Name IN ["Weather", "InstantNetSpeed"]')
  .executeInParallel('compileKeywordCheckerReport')
}

function compileKeywordCheckerReport() {
  var accName = AdWordsApp.currentAccount().getName();
  var ss = SpreadsheetApp.openByUrl(KEYWORD_CHECKER_URL);
  var tab = ss.getSheetByName(accName);
  if(!tab) { 
    return;
  }
  
  var data = tab.getDataRange().getValues();
  var header = data.shift();
  
  var dt = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'dd'), 10);
  //if(header[1] && header[1].getDate() == dt) {
    //return;
  //}
  
  data.shift();
  data.shift();
  
  if(!data.length) { return; }
  
  var map = {}
  var OPTIONS = { 'includeZeroImpressions' : true };
  
  var cols = ['AdGroupId','Criteria','KeywordMatchType', 'Status'/*,'IsNegative'*/];
  var report = 'KEYWORDS_PERFORMANCE_REPORT';
  var query = ['select',cols.join(','),'from',report,
               'where AdGroupStatus = ENABLED and CampaignStatus = ENABLED',
               'and Status IN [ENABLED,PAUSED]',
               'during','YESTERDAY'].join(' ');
  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while(reportIter.hasNext()) {
    var row = reportIter.next();
    var cleanKeyword = row.Criteria.replace(/[+]/g,'').toLowerCase();
    
    if(!map[cleanKeyword]) {
      map[cleanKeyword] = {
        'Exact': '', 'BMM': '', 'Phrase': '', 'Broad': ''
      }
    }
    
    var status = 'E';
    if(row.Status == 'paused') {
      status = 'P';  
    }
    
    if(row.KeywordMatchType == 'Broad') {
      if(row.Criteria.indexOf('+') > -1) {
        if(map[cleanKeyword]['BMM'] == 'E') { continue; }
        map[cleanKeyword]['BMM'] = status;
      } else {
        if(map[cleanKeyword]['Broad'] == 'E') { continue; }
        map[cleanKeyword]['Broad'] = status;
      }
    } else {
      if(map[cleanKeyword][row.KeywordMatchType] == 'E') { continue; }   
      map[cleanKeyword][row.KeywordMatchType] = status;
    }    
  }
  
  for(var z in data) {
    if(!data[z][0]) { continue; }
    var kw = data[z][0].toLowerCase();
    if(!map[kw]) {
      map[kw] = {
        'Exact': '', 'BMM': '', 'Phrase': '', 'Broad': ''
      }
    }  
    
    data[z][1] = map[kw]['Exact'];
    data[z][2] = map[kw]['BMM'];
    data[z][3] = map[kw]['Phrase'];
    data[z][4] = map[kw]['Broad'];
  }
  
  tab.getRange(4, 1, data.length, data[0].length).setValues(data);  
  tab.getRange(1,2).setValue(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), "MMM d, yyyy HH:mm"));
}