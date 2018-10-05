function main() {
    var map = getApprovedProducts();
    //Logger.log(Object.keys(map).length);
  
    var url = 'https://docs.google.com/spreadsheets/d/1dEGQkrOdVr22UDohi2t4tBxy6hziUQvoUkwRmGG9dh4/edit';
    var tab = SpreadsheetApp.openByUrl(url).getSheets()[0];
    var data = tab.getDataRange().getValues();
    data.shift();
    
    for(var z in data) {
      //if(!data[z][25]) { continue; }
      
      if(map[data[z][0]]) {
        data[z][28] = 'Yes'; 
      } else {
        data[z][28] = 'No'; 
      }
    }
    
    tab.getRange(2,1,data.length,data[0].length).setValues(data);
  }
  
  function getApprovedProducts() {
    var map = {};
    
    var MERCHANT_ID = 114554920;
    var pageToken;
    var pageNum = 1;
    var maxResults = 250;
    
      // List all the products for a given merchant.
    do {
      var products = ShoppingContent.Productstatuses.list(MERCHANT_ID, {
        pageToken: pageToken,
        maxResults: maxResults
      });
      
      for(var i in products.resources) {
        for(var k in products.resources[i]['destinationStatuses']) {
          var status =  products.resources[i]['destinationStatuses'][k];
          if(status["destination"] == "Shopping" && status["intention"] != "excluded" && status["approvalStatus"] == "approved") {
            var id = parseInt(products.resources[i]['productId'].replace('online:en:GB:', ''), 10);
            map[id] = 1;
          }
        }
      } 
      pageToken = products.nextPageToken;
      pageNum++;
    } while (pageToken);
  
    return map; 
  }