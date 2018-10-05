/*************************************************
* Ad Customizer Manager
* @author Naman Jindal <naman@pushgroup.co.uk>
* @version 1.0
***************************************************/

var URL = 'https://docs.google.com/spreadsheets/d/1yWIjfM9nKlTg-GybIuvy2I8Paf8awdxh5mySINAhN18/edit#gid=1000431829';

function main(){
  MccApp.accounts().withIds(['665-818-8003']).executeInParallel('run');
}

function run() {
  var ignoreList = ['start date','end date','device preference','scheduling','target campaign','target ad group',
                    'target adgroup','target keyword','path 1','path 2', 'final url',
                    'keyword text','match type','keyword','id'];
  var SETTINGS = new Object();  
  SETTINGS.URL = URL;
  
  pullBestAdsToSpreadsheet(SETTINGS);
  cleanupPausedAdGroups(SETTINGS);
  
  var inputData = SpreadsheetApp.openByUrl(SETTINGS.URL).getSheetByName(AdWordsApp.currentAccount().getName()).getDataRange().getValues();
  if(inputData.length < 3) { info('No Data'); return; }
  
  SETTINGS.DATASOURCENAME = inputData[0][1];
  if(!SETTINGS.DATASOURCENAME) {
    SETTINGS.DATASOURCENAME = 'PushAds';
  }
  
  var source = getOrCreateDataSource(SETTINGS,inputData);
  
  var customizers = source.items().get();
  var customizersById = {};
  while (customizers.hasNext()) {
    var customizer = customizers.next();
    var ag = customizer.getTargetAdGroupName();
    //var kw = customizer.getTargetKeywordText();
    var camp = customizer.getTargetCampaignName();
    if(!ag) { ag = ''; }
    //if(!kw) { kw = ''; }
    customizersById[[camp,ag].join('~~')] = customizer;
  }
  
  var formatHeader = inputData.shift();
  var header = inputData.shift();
  for(var k in inputData) {
    var customizer = customizersById[[inputData[k][0],inputData[k][1]].join('~~')];
    if(customizer) {
      //Logger.log(customizer.getAttributeValue('CountDownDate'));
      for(var j in header) {
        if(ignoreList.indexOf(header[j].toLowerCase()) > -1) { continue; }
        var value = inputData[k][j];
        if(formatHeader[j] == 'date') {
          value = value.replace(/[/]/g,'').replace(/:/g,'');
        }
        customizer.setAttributeValue(header[j], value);
      }	
    } else {
      var item = source.adCustomizerItemBuilder()
      for(var j in header) {
        if(ignoreList.indexOf(header[j].toLowerCase()) > -1) { continue; }
        var value = inputData[k][j];
        if(formatHeader[j] == 'date') {
          value = value.replace(/[/]/g,'').replace(/:/g,'');
        }
        item.withAttributeValue(header[j], value);
      }	
      
      if(inputData[k][1]) {
        item.withTargetAdGroup(inputData[k][0],inputData[k][1]).build();
      } else {
        item.withTargetCampaign(inputData[k][0]).build();	  
      }  
    }
  }
  
  var pause_a = {}, pause_b = {}, pause_c = {};
  var customizers = source.items().get();
  var customizersById = {};
  while (customizers.hasNext()) {
    var customizer = customizers.next();
    var ag = customizer.getTargetAdGroupName();
    var camp = customizer.getTargetCampaignName();
    if(!ag) { ag = ''; }
    
    if(!customizer.getAttributeValue('HL1a')) {
      pause_a[[camp,ag].join('~~')] = 1;
    } 
    
    if(!customizer.getAttributeValue('HL1b')) {
      pause_b[[camp,ag].join('~~')] = 1;
    }
    
    if(!customizer.getAttributeValue('HL1c')) {
      pause_c[[camp,ag].join('~~')] = 1;
    }
  }
  
  var iter = AdWordsApp.ads()
  .withCondition('Status = ENABLED')
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL"')
  .get();
  
  while(iter.hasNext()) { 
    var ad = iter.next().asType().expandedTextAd();
    var key = [ad.getCampaign().getName(), ad.getAdGroup().getName()].join('~~');
    if(ad.getHeadlinePart1().indexOf('PushAds.HL1a') > -1) {
      if(pause_a[key]) {
        ad.pause(); 
      }
    } else if(ad.getHeadlinePart1().indexOf('PushAds.HL1b') > -1) {
      if(pause_b[key]) {
        ad.pause(); 
      }
    } else if(ad.getHeadlinePart1().indexOf('PushAds.HL1c') > -1) {
      if(pause_c[key]) {
        ad.pause(); 
      }
    }
  }
}  

function getOrCreateDataSource(SETTINGS,sourceData) {
  var sources = AdWordsApp.adCustomizerSources().get();
  while (sources.hasNext()) {
    var source = sources.next();
    if (source.getName() == SETTINGS.DATASOURCENAME) {
      return source;
    }
  }
  
  var source = AdWordsApp.newAdCustomizerSourceBuilder().withName(SETTINGS.DATASOURCENAME)
  
  for(var k in sourceData[1]) {
    if(ignoreList.indexOf(sourceData[1][k].toLowerCase()) > -1) { continue; }
    source = source.addAttribute(sourceData[1][k], sourceData[0][k] ? sourceData[0][k].toLowerCase() : 'text');
  }
  
  return source.build().getResult();    
}

function activeCustomizerAds() {
  var ids = [];
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = DISPLAY').get();
  while(iter.hasNext()) {
    ids.push(iter.next().getId());
  }
  
  setupLabel('Customizer Ad 1');
  setupLabel('Customizer Ad 2');
  setupLabel('Customizer Ad 3');
  
  var iter = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL1a"')
  .withCondition('Status = PAUSED')
  .withCondition('LabelNames CONTAINS_NONE ["Customizer Ad 1"]')
  
  if(ids.length) {    
    iter.withCondition('CampaignId NOT_IN [' + ids.join(',') + ']')
  }
  
  iter = iter.withCondition('CampaignStatus IN [ENABLED,PAUSED]')
  .withCondition('AdGroupStatus IN [ENABLED,PAUSED]')
  .get();
  
  while(iter.hasNext()) {
    var ad = iter.next();
    ad.enable();   
    ad.applyLabel('Customizer Ad 1');
  }
  
  var iter = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL1b"')
  .withCondition('Status = PAUSED')
  .withCondition('LabelNames CONTAINS_NONE ["Customizer Ad 2"]')
  
  if(ids.length) {    
    iter.withCondition('CampaignId NOT_IN [' + ids.join(',') + ']')
  }
  
  iter = iter.withCondition('CampaignStatus IN [ENABLED,PAUSED]')
  .withCondition('AdGroupStatus IN [ENABLED,PAUSED]')
  .get();
  
  while(iter.hasNext()) {
    var ad = iter.next();
    ad.enable();   
    ad.applyLabel('Customizer Ad 2');
  }
  
  var iter = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL1c"')
  .withCondition('Status = PAUSED')
  .withCondition('LabelNames CONTAINS_NONE ["Customizer Ad 3"]')
  
  if(ids.length) {    
    iter.withCondition('CampaignId NOT_IN [' + ids.join(',') + ']')
  }
  
  iter = iter.withCondition('CampaignStatus IN [ENABLED,PAUSED]')
  .withCondition('AdGroupStatus IN [ENABLED,PAUSED]')
  .get();
  
  while(iter.hasNext()) {
    var ad = iter.next();
    ad.enable();   
    ad.applyLabel('Customizer Ad 3');
  }
}

function setupLabel(name) {
  if(!AdWordsApp.labels().withCondition('Name = "'+ name + '"').get().hasNext()) {
    AdWordsApp.createLabel(name);    
  }
}

function pullBestAdsToSpreadsheet(SETTINGS) {
  
  if(!AdWordsApp.labels().withCondition('Name = "Top ETA 1"').get().hasNext()) {
    AdWordsApp.createLabel('Top ETA 1');
  }
  
  var settingsMap = {}, agMap = {};
  var ags = AdWordsApp.adGroups()
  .withCondition('Status IN [ENABLED]')
  .withCondition('CampaignStatus IN [ENABLED]')
  .get();
  while(ags.hasNext()) {
    var ag = ags.next();
    var key = [ag.getCampaign().getName(), ag.getName()].join('!~!');
    agMap[key] = ag.getId();
  }
  
  var map = {};
  var ads = AdWordsApp.ads()
  .withCondition('LabelNames CONTAINS_ANY ["Top ETA 1"]')
  .withCondition('Status IN [ENABLED]')
  .withCondition('CampaignStatus IN [ENABLED]')
  .withCondition('AdGroupStatus IN [ENABLED]')
  .get();
  while(ads.hasNext()) {
    var ad = ads.next().asType().expandedTextAd();  
    var key = [ad.getCampaign().getName(), ad.getAdGroup().getName()].join('!~!');
    map[key] = [ad.getCampaign().getName(), ad.getAdGroup().getName(), '',
                ad.getHeadlinePart1(), ad.getHeadlinePart2(), ad.getDescription(),
                ad.getHeadlinePart1(), ad.getHeadlinePart2(), ad.getDescription()];
    
    settingsMap[ad.getAdGroup().getId()] = {
      'FinalUrl': ad.urls().getFinalUrl(), 'Path1': ad.getPath1(), 'Path2': ad.getPath1()
    }                 
    delete agMap[key];
  } 
  
  var ids = [];
  for(var key in agMap) {
    ids.push(agMap[key])
  }
  
  if(ids.length) {
    var ads = AdWordsApp.ads()
    .withCondition('HeadlinePart1 DOES_NOT_CONTAIN_IGNORE_CASE "PushAds"')
    .withCondition('Status IN [ENABLED]')
    .withCondition('Type = EXPANDED_TEXT_AD')
    .withCondition('AdGroupId IN ['+ids.join(',')+']')
    .orderBy('Conversions DESC')
    .orderBy('Clicks DESC')
    .forDateRange('LAST_30_DAYS')
    .get();
    
    while(ads.hasNext()) {
      var ad = ads.next().asType().expandedTextAd();  
      var key = [ad.getCampaign().getName(), ad.getAdGroup().getName()].join('!~!');
      map[key] = [ad.getCampaign().getName(), ad.getAdGroup().getName(), '',
                  ad.getHeadlinePart1(), ad.getHeadlinePart2(), ad.getDescription(),
                  ad.getHeadlinePart1(), ad.getHeadlinePart2(), ad.getDescription()];
      settingsMap[ad.getAdGroup().getId()] = {
        'FinalUrl': ad.urls().getFinalUrl(), 'Path1': ad.getPath1(), 'Path2': ad.getPath1()
      }                   
    } 
  }
  
  //info(Object.keys(map).length);
  
  var AC_TEMPLATE_URL = 'https://docs.google.com/spreadsheets/d/17_XJWz4Wwwe-zIHYn4oihAuJgMpkIJSEXdHTgbCJe3o/edit#gid=739633464'
  
  var rep = SpreadsheetApp.openByUrl(SETTINGS.URL);
  var tab = rep.getSheetByName(accName);
  if(!tab) {
    tab = SpreadsheetApp.openByUrl(AC_TEMPLATE_URL).getSheets()[0].copyTo(rep);
    tab.setName(accName);
  }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  
  for(var z in data) {
    var key = [data[z][0],data[z][1]].join('!~!');
    if(map[key]) {
      delete map[key];
    } 
  }
  
  //tab.getRange(3,1,data.length,data[0].length).setValues(data);
  
  var out = [];
  for(var key in map) {
    out.push(map[key]);
  }
  
  if(out.length) {
    tab.getRange(tab.getLastRow()+1, 1, out.length, out[0].length).setValues(out);
  }  
  
  uploadMissingCustomizerAds(settingsMap);
  
  deleteRemovedCampaignRows(tab);
}

function deleteRemovedCampaignRows(tab) {
  var camps = {};
  var iter = AdWordsApp.campaigns().withCondition('Status = REMOVED').get();
  while(iter.hasNext()) {
    camps[iter.next().getName()] = 1;    
  }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var toRemove = [];
  for(var z in data) {
    if(camps[data[z][0]]) {
      toRemove.push(parseInt(z,10)+3);  
    } 
  }
  
  for(var z = toRemove.length-1; z >= 0; z--) {
    tab.deleteRow(toRemove[z]);
  }
}

function uploadMissingCustomizerAds(settingsMap) {
  //info('here');
  
  
  var tempMap = JSON.parse(JSON.stringify(settingsMap))  ;
  var adMap = {};  
  var ads = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL1a"')
  .orderBy('Conversions DESC')
  .orderBy('Clicks DESC')
  .forDateRange('THIS_MONTH')
  .get();
  
  while(ads.hasNext()) {
    var ad = ads.next();
    var key = [ad.getAdGroup().getId(), ad.getHeadlinePart1()].join('!~!');
    if(adMap[key]) { 
      ad.remove();
      continue;
    }
    
    adMap[key] = 1;
    
    var id = ad.getAdGroup().getId();
    delete tempMap[id];
  }
  
  var ids = Object.keys(tempMap);
  if(ids.length) { 
    var columnHeads = ['Ad group ID', 'Headline 1', 'Headline 2', 'Description',
                       'Path 1', 'Path 2', 'Final URL', 'Ad state'];
    var upload = AdWordsApp.bulkUploads().newCsvUpload(columnHeads, {moneyInMicros: false});
    
    for(var id in tempMap) {
      upload.append({
        'Ad group ID': id, 'Headline 1': '{=PushAds.HL1a}', 'Headline 2': '{=PushAds.HL2a}', 
        'Description': '{=PushAds.DLa}', 'Path 1': tempMap[id].Path1, 'Path 2': tempMap[id].Path2, 
        'Final URL': tempMap[id].FinalUrl, 'Ad state': 'paused'
      });
      
      upload.append({
        'Ad group ID': id, 'Headline 1': '{=PushAds.HL1b}', 'Headline 2': '{=PushAds.HL2b}', 
        'Description': '{=PushAds.DLb}', 'Path 1': tempMap[id].Path1, 'Path 2': tempMap[id].Path2, 
        'Final URL': tempMap[id].FinalUrl, 'Ad state': 'paused'
      });
      
      upload.append({
        'Ad group ID': id, 'Headline 1': '{=PushAds.HL1c}', 'Headline 2': '{=PushAds.HL2c}', 
        'Description': '{=PushAds.DLc}', 'Path 1': tempMap[id].Path1, 'Path 2': tempMap[id].Path2, 
        'Final URL': tempMap[id].FinalUrl, 'Ad state': 'paused'
      });
    }
    
    upload.apply();
  }
  
  var tempMap = JSON.parse(JSON.stringify(settingsMap))  ;
  var adMap = {};  
  var ads = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL1c"')
  .orderBy('Conversions DESC')
  .orderBy('Clicks DESC')
  .forDateRange('THIS_MONTH')
  .get();
  
  while(ads.hasNext()) {
    var ad = ads.next();
    var key = [ad.getAdGroup().getId(), ad.getHeadlinePart1()].join('!~!');
    if(adMap[key]) { 
      ad.remove();
      continue;
    }
    
    adMap[key] = 1;
    
    var id = ad.getAdGroup().getId();
    delete tempMap[id];
  }
  
  var ids = Object.keys(tempMap);
  if(!ids.length) { info('No IDS'); return; }
  
  var columnHeads = ['Ad group ID', 'Headline 1', 'Headline 2', 'Description',
                     'Path 1', 'Path 2', 'Final URL', 'Ad state'];
  var upload = AdWordsApp.bulkUploads().newCsvUpload(columnHeads, {moneyInMicros: false});
  
  for(var id in tempMap) {
    upload.append({
      'Ad group ID': id, 'Headline 1': '{=PushAds.HL1c}', 'Headline 2': '{=PushAds.HL2c}', 
      'Description': '{=PushAds.DLc}', 'Path 1': tempMap[id].Path1, 'Path 2': tempMap[id].Path2, 
      'Final URL': tempMap[id].FinalUrl, 'Ad state': 'paused'
    });
  }
  
  upload.apply();
  
}

function deActivateCustomizerAds() {
  var iter = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL"')
  .withCondition('Status = ENABLED')
  .get();
  
  while(iter.hasNext()) {
    iter.next().pause();    
  }  
}

function deactivateDisplayCustomizerAds() {
  
  var ids = [];
  var iter = AdWordsApp.campaigns().withCondition('AdvertisingChannelType = DISPLAY').get();
  while(iter.hasNext()) {
    ids.push(iter.next().getId());
  }
  if(!ids.length) { return; }
  
  var iter = AdWordsApp.ads()
  .withCondition('HeadlinePart1 CONTAINS "PushAds.HL"')
  .withCondition('CampaignId IN [' + ids.join(',') + ']')
  .withCondition('Status = ENABLED')
  .get();
  
  while(iter.hasNext()) {
    iter.next().pause();    
  }
}

function cleanupPausedAdGroups(SETTINGS) {
  
  var iter = AdWordsApp.adGroups()
  .withCondition('Status = PAUSED')   
  .withCondition('CampaignStatus = ENABLED')   
  .get();
  
  var map = {};
  while(iter.hasNext()) {
    var ag = iter.next();
    var key = [ag.getCampaign().getName(), ag.getName()].join('!~!');
    map[key] = 1;
  }
  
  var rep = SpreadsheetApp.openByUrl(SETTINGS.URL);
  var tab = rep.getSheetByName(accName);
  if(!tab) { return; }
  
  var data = tab.getDataRange().getValues();
  data.shift();
  data.shift();
  
  var len = data.length - 1;
  for(var z = len; z >= 0; z--) {
    var key = [data[z][0],data[z][1]].join('!~!');
    if(map[key]) {
      tab.deleteRow(parseInt(z,10)+3);
    } 
  }
}