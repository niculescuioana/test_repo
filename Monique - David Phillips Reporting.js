// Monthly on 1st at 12 PM

var FOLDER_ID = '1ll26Zpsz4kJlawAiLqCNXZOMy_TGr4Bq';

function main() {
  MccApp.accounts().withCondition('Name = "David Phillips"').executeInParallel('run');
}

function run() {
  
  var date = new Date(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'MMM d, yyyy'));
  date.setHours(12);
  date.setDate(1);
  date.setDate(0);
  
  var month = Utilities.formatDate(date, 'PST', 'MMM yyyy');
  
  var ss = SpreadsheetApp.create('David Phillips: Campaign Performance (' + month + ')');
  var campTab = ss.getActiveSheet();
  campTab.setName('Campaigns');
  var adTab = ss.insertSheet('Ads'),
      kwTab = ss.insertSheet('Keywords');
  
  var header = [
    ['Campaign', 'Budget',  'Clicks', 'Impressions', 'CTR', 'Avg CPC', 'Cost', 'Avg Pos',
     'Conversions', 'Cost per conv', 'Conv rate', 'Search Impr. share', 'Search Lost IS (rank)', 'Search Lost IS (budget)', 'Labels']
  ];
  
  var query = [
    'SELECT CampaignName, Amount, Clicks, Impressions, Ctr, AverageCpc, Cost, AveragePosition, Conversions, CostPerConversion,',
    'ConversionRate, SearchImpressionShare, SearchRankLostImpressionShare, SearchBudgetLostImpressionShare, Labels',
    'FROM CAMPAIGN_PERFORMANCE_REPORT DURING LAST_MONTH'
  ].join(' ');
  
  AdWordsApp.report(query, { 'includeZeroImpressions': false }).exportToSheet(campTab);
  campTab.getRange(1,1,1,header.length).setValues([header]).setFontWeight('Bold');
  campTab.getRange(1,1,1,header.length)
  campTab.getDataRange().setFontFamily('Calibri');
  
  var header = ['Campaign', 'Ad Group', 'Labels', 'Keyword', 'Match Type', 'Clicks', 'Impressions', 'CTR', 'Avg CPC', 'Cost', 'Avg Pos',
                'Conversions', 'Cost per conv', 'Conv rate'];
  var query = [
    'SELECT CampaignName, AdGroupName, Labels, Criteria, KeywordMatchType,',
    'Clicks, Impressions, Ctr, AverageCpc, Cost, AveragePosition, Conversions, CostPerConversion, ConversionRate',
    'FROM KEYWORDS_PERFORMANCE_REPORT DURING LAST_MONTH'
  ].join(' ');
  
  AdWordsApp.report(query, { 'includeZeroImpressions': false }).exportToSheet(kwTab);
  
  kwTab.getRange(1,1,1,header.length).setValues([header]).setFontWeight('Bold');
  kwTab.getDataRange().setFontFamily('Calibri');
  
  var out = [
    ['Campaign', 'Ad Group', 'Labels', 'Headline 1', 'Headline 2', 'Description', 'Path 1', 'Path 2', 'Final Url',
     'Clicks', 'Impressions', 'CTR', 'Avg CPC', 'Cost', 'Avg Pos',
     'Conversions', 'Cost per conv', 'Conv rate']
  ];
  
  var query = [
    'SELECT CampaignName, AdGroupName, Labels, HeadlinePart1, HeadlinePart2, Description, Path1, Path2, CreativeFinalUrls,',
    'Clicks, Impressions, Ctr, AverageCpc, Cost, AveragePosition, Conversions, CostPerConversion, ConversionRate',
    'FROM AD_PERFORMANCE_REPORT WHERE AdType = EXPANDED_TEXT_AD DURING LAST_MONTH'
  ].join(' ');
  
  var rows = AdWordsApp.report(query, { 'includeZeroImpressions': false }).rows();
  while(rows.hasNext()) {
    var row = rows.next();
    row.CreativeFinalUrls = JSON.parse(row.CreativeFinalUrls)[0];
    out.push([row.CampaignName, row.AdGroupName, row.Labels, row.HeadlinePart1, row.HeadlinePart2, row.Description, row.Path1, row.Path2, row.CreativeFinalUrls,
              row.Clicks, row.Impressions, row.Ctr, row.AverageCpc, row.Cost, row.AveragePosition, row.Conversions, row.CostPerConversion, row.ConversionRate]);
  }
  
  adTab.setFrozenRows(1);
  adTab.getRange(1,1,1,out[0].length).setFontWeight('Bold');
  adTab.getRange(1,1,out.length,out[0].length).setValues(out).setFontFamily('Calibri');
  
  ss.addEditor('analytics@pushgroup.co.uk');
  var folder = DriveApp.getFolderById(FOLDER_ID);
  folder.addFile(DriveApp.getFileById(ss.getId()));
}
