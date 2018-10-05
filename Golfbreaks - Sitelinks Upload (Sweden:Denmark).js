function main() {
    //MccApp.accounts().withCondition('Name = "Golfbreaks - Sweden"').executeInParallel('run');
    //MccApp.accounts().withCondition('Name = "Golfbreaks - Denmark"').executeInParallel('cleanDenmark');
    MccApp.accounts().withCondition('Name = "Golfbreaks"').executeInParallel('cleanUK');
  }
  
  function cleanUK() {
    
    var cidMap = {
      "Brand": 999733961,
      "Competitors": 999733962,
      "Generic": 999733963
    }
    
    var slMap = {};
    var iter = AdWordsApp.extensions().sitelinks().get();
    
    while(iter.hasNext()) {
      var sl = iter.next();
      var url = sl.urls().getFinalUrl();
      if(url.indexOf('cid=999733961') > -1) {
        slMap['Brand'] = sl;
      } else if(url.indexOf('cid=999733962') > -1) {
        slMap['Competitors'] = sl;
      } else if(url.indexOf('cid=999733963') > -1) {
        slMap['Generic'] = sl;
      }
    }
    
    
    Logger.log(JSON.stringify(slMap))
    
  }
  
  function run() {
    var linkText = "Prenumerera på nyhetsbrev"
    var desc1 = "Prenumerera på våra nyhetsbrev för";
    var desc2 = "att få exklusiva specialerbjudanden";
    
  /*  var linkText = "Tilmelding til nyhedsbrev"
    var desc1 = "Tilmeld dig vores nyhedsbrev i dag";
    var desc2 = "og modtag vores ekslusive tilbud";
    */
    
    var linkUrl = 'http://www.golfbreaks.com/se/nyhetsbrev/?utm_source=google&utm_medium=cpc&utm_term={keyword}&utm_campaign=[CAMPAIGN_NAME]';
    
    var slByUrl = {}, newSitelink = {};
    
    var iter = AdWordsApp.extensions().sitelinks().get();
    while(iter.hasNext()) {
      var entity = iter.next();
      if(entity.getLinkText() == linkText && entity.getDescription1() == desc1 && entity.getDescription2() == desc2) {
        var url = entity.urls().getFinalUrl();
        slByUrl[url] = entity;
      }
    }
    
    var adGroups = AdWordsApp.adGroups()
    .withCondition('Status = ENABLED')
    .withCondition('CampaignStatus = ENABLED')
    .get();
    
    while(adGroups.hasNext()) {
      var ag = adGroups.next();
      
      var campName = ag.getCampaign().getName();
      var skipAg = false;
      
      var finalUrl = encodeURI(linkUrl.replace('[CAMPAIGN_NAME]', campName));
      var campaignSitelink = slByUrl[finalUrl];
      if(!campaignSitelink) {
        campaignSitelink = AdWordsApp.extensions().newSitelinkBuilder()
        .withLinkText(linkText)
        .withDescription1(desc1)
        .withDescription2(desc2)
        .withFinalUrl(finalUrl)
        .build()
        .getResult();
        newSitelink[finalUrl] = 1;
      }
      
      if(!newSitelink[finalUrl] || !campaignSitelink) { 
        //Logger.log(finalUrl);
        continue; 
      }
      slByUrl[finalUrl] = campaignSitelink;
      
      /*var sls = ag.extensions().sitelinks().get();
      while(sls.hasNext()) {
        var sl = sls.next();
        if(sl.getLinkText() == linkText && sl.getDescription1() == desc1 
        && sl.getDescription2() == desc2 && sl.urls().getFinalUrl() == finalUrl) {
          skipAg = true;
          break;
        }
      }
      
      if(skipAg) { continue; }*/
      
      ag.addSitelink(campaignSitelink);
    }
  }