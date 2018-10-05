function main() {
    //MccApp.accounts().withCondition('Name = "Golfbreaks"').executeInParallel('test') ;
     MccApp.accounts().withCondition('Name = "Golfbreaks - USA Account"').executeInParallel('run') ;
   }
   
   function test() {
     var kws = AdWordsApp.keywords()
     .withCondition('FinalUrls CONTAINS "/cid="')
     .withCondition('CampaignStatus IN [ENABLED,PAUSED]')
     .withCondition('AdGroupStatus IN [ENABLED,PAUSED]')
     .withCondition('Status IN [ENABLED,PAUSED]')    
     .get();
     while(kws.hasNext()) {
       var kw = kws.next();
       var url = kw.urls().getFinalUrl().replace('/cid=', '?cid=');
       kws.next().urls().setFinalUrl(url);
     }
   }
   
   function run() {
     var url = 'https://docs.google.com/spreadsheets/d/1ILAj5tn68UllKaeKmSsVGrf6-n-pr-R2ONSBKX9DxTs/edit#gid=889106203';
     var tab = SpreadsheetApp.openByUrl(url).getSheetByName('Copy of '+AdWordsApp.currentAccount().getName());
     var data = tab.getDataRange().getValues();
     data.shift();
     
     for(var x in data) {
       if(!data[x][4] || !data[x][3]) { continue; }
       
       var kws = AdWordsApp.keywords()
       .withCondition('FinalUrls = "'+data[x][3]+'"')
       .withCondition('CampaignStatus IN [ENABLED,PAUSED]')
       .withCondition('AdGroupStatus IN [ENABLED,PAUSED]')
       .withCondition('Status IN [ENABLED,PAUSED]')    
       .get();
       while(kws.hasNext()) {
         kws.next().urls().setFinalUrl(data[x][4]);
       }
     }
   }