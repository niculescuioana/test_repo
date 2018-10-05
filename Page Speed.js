var API_KEY = 'AIzaSyDMLaoRpd_0kBJEqKCz83oDneh1g_L068U';
var PAGESPEED_URL =
    'https://www.googleapis.com/pagespeedonline/v2/runPagespeed?';


/**
* Specifies the amount of time in seconds required at the end of each parallel
* execution to collect up the results and return them.
*/
var POST_ITERATION_TIME_SECS = 5 * 60;

/**
* The actual limit on the number of URLs to fetch from a single account — this
* is only supposed to be a sample of the performance of URLs in this account,
* not exhaustive.
*/
var URL_LIMIT = 250;

/**
* Specified the amount of time after the URL fetches required to write to and
* format the spreadsheet.
*/
var SPREADSHEET_PREP_TIME_SECS = 4 * 60;

/**
* Represents the number of retries to use with the PageSpeed service.
*/
var MAX_RETRIES = 3;

function main() {
  //SpreadsheetApp.openByUrl(url);
  var REPORT_URL = 'https://docs.google.com/spreadsheets/d/1uFytKncJ6wbSl2pxvElRBu4CkwOkSvvLfAExxYs6OXE/edit';
  var tab = SpreadsheetApp.openByUrl(REPORT_URL).getSheetByName('Test')
  var data = tab.getDataRange().getValues();
  
  data.shift();
  data.shift();
  data.shift();
  
  for(var x in data) {
    var row = parseInt(x,10)+4;
    if(data[x][1] && data[x][1] != 'N/A' && data[x][2] === '') {
      var result = checkUrl(data[x][1]);
      if(result) {
        tab.getRange(row, 3, 1, 2).setValues(result);
      }
    }
    
    if(data[x][4] && data[x][4] != 'N/A' && data[x][5] === '') {
      var result = checkUrl(data[x][4]);
      if(result) {
        tab.getRange(row, 6, 1, 2).setValues(result);
      }
    }
    
    if(data[x][7] && data[x][7] != 'N/A' && data[x][8] === '') {
      var result = checkUrl(data[x][7]);
      if(result) {
        tab.getRange(row, 9, 1, 2).setValues(result);
      }
    }
    
    if(data[x][10] && data[x][10] != 'N/A' && data[x][11] === '') {
      var result = checkUrl(data[x][10]);
      if(result) {
        tab.getRange(row, 12, 1, 2).setValues(result);
      }
    }
  }
}


function getPageSpeedResultsForUrls(urlStore) {
  // Associative array for column headings and contextual help URLs.
  var headings = {};
  var errors = {};
  // Record results on a per-URL basis.
  var pageSpeedResults = {};
  
  for (var url in urlStore) {
    if (hasRemainingTimeForUrlFetches()) {
      var result = getPageSpeedResultForSingleUrl(url);
      if (!result.error) {
        pageSpeedResults[url] = result.pageSpeedInfo;
        var columnsResult = result.columnsInfo;
        // Loop through each heading element; the PageSpeed Insights API
        // doesn't always return URLs for each column heading, so aggregate
        // these across each call to get the most complete list.
        var columnHeadings = Object.keys(columnsResult);
        for (var i = 0, lenI = columnHeadings.length; i < lenI; i++) {
          var columnHeading = columnHeadings[i];
          if (!headings[columnHeading] || (headings[columnHeading] &&
              headings[columnHeading].length <
              columnsResult[columnHeading].length)) {
            headings[columnHeading] = columnsResult[columnHeading];
          }
        }
      } else {
        errors[url] = result.error;
      }
    }
  }
  
  var tableHeadings = ['URL', 'Speed', 'Usability'];
  var headingKeys = Object.keys(headings);
  for (var y = 0, lenY = headingKeys.length; y < lenY; y++) {
    tableHeadings.push(headingKeys[y]);
  }
  
  var table = [];
  var pageSpeedResultsUrls = Object.keys(pageSpeedResults);
  for (var r = 0, lenR = pageSpeedResultsUrls.length; r < lenR; r++) {
    var resultUrl = pageSpeedResultsUrls[r];
    var row = [toPageSpeedHyperlinkFormula(resultUrl)];
    var data = pageSpeedResults[resultUrl];
    for (var j = 1, lenJ = tableHeadings.length; j < lenJ; j++) {
      row.push(data[tableHeadings[j]]);
    }
    table.push(row);
  }
  // Present the table back in the order worst-performing-first.
  table.sort(function(first, second) {
    if (first[1] + first[2] < second[1] + second[2]) {
      return -1;
    } else if (first[1] + first[2] > second[1] + second[2]) {
      return 1;
    }
    return 0;
  });
  
  // Add hyperlinks to all column headings where they are available.
  for (var h = 0, lenH = tableHeadings.length; h < lenH; h++) {
    // Sheets cannot have multiple links in a single cell at the moment :-/
    if (headings[tableHeadings[h]] &&
        typeof(headings[tableHeadings[h]]) === 'object') {
      tableHeadings[h] = '=HYPERLINK("' + headings[tableHeadings[h]][0] +
        '","' + tableHeadings[h] + '")';
    }
  }
  table.unshift(tableHeadings);
  
  // Form table from errors
  var errorTable = [];
  var errorKeys = Object.keys(errors);
  for (var k = 0; k < errorKeys.length; k++) {
    errorTable.push([errorKeys[k], errors[errorKeys[k]]]);
  }
  return {
    table: table,
    errors: errorTable
  };
}

/**
* Given a URL, returns a spreadsheet formula that displays the URL yet links to
* the PageSpeed URL for examining this.
*
* @param {String} url The URL to embed in the Hyperlink formula.
* @return {String} A string representation of the spreadsheet formula.
*/
function toPageSpeedHyperlinkFormula(url) {
  return '=HYPERLINK("' +
    'https://developers.google.com/speed/pagespeed/insights/?url=' + url +
      '&tab=mobile","' + url + '")';
}

/**
* Creates an object of results metrics from the parsed results of a call to
* the PageSpeed service.
*
* @param {Object} parsedPageSpeedResponse The object returned from PageSpeed.
* @return {Object} An associative array with entries for each metric.
*/
function extractResultRow(parsedPageSpeedResponse) {
  var urlScores = {};
  if (parsedPageSpeedResponse.ruleGroups) {
    var ruleGroups = parsedPageSpeedResponse.ruleGroups;
    var usabilityScore = ruleGroups.USABILITY.score;
    var speedScore = ruleGroups.SPEED.score;
    urlScores.Speed = speedScore;
    urlScores.Usability = usabilityScore;
  }
  if (parsedPageSpeedResponse.formattedResults &&
      parsedPageSpeedResponse.formattedResults.ruleResults) {
    var resultParts = parsedPageSpeedResponse.formattedResults.ruleResults;
    for (var partName in resultParts) {
      var part = resultParts[partName];
      urlScores[part.localizedRuleName] = part.ruleImpact;
    }
  }
  return urlScores;
}

/**
* Extracts the headings for the metrics returned from PageSpeed, and any
* associated help URLs.
*
* @param {Object} parsedPageSpeedResponse The object returned from PageSpeed.
* @return {Object} An associative array used to store column-headings seen
*     in the response. This can take two forms:
*     (1) {'heading':'heading', ...} - this form is where no help URLs are
*     known.
*     (2) {'heading': [url1, ...]} - where one or more URLs is returned that
*     provides help on the particular heading item.
*/
function extractColumnsInfo(parsedPageSpeedResponse) {
  var columnsInfo = {};
  if (parsedPageSpeedResponse.formattedResults &&
      parsedPageSpeedResponse.formattedResults.ruleResults) {
    var resultParts = parsedPageSpeedResponse.formattedResults.ruleResults;
    for (var partName in resultParts) {
      var part = resultParts[partName];
      if (!columnsInfo[part.localizedRuleName]) {
        columnsInfo[part.localizedRuleName] = part.localizedRuleName;
      }
      // Find help URLs in the response.
      var summary = part.summary;
      if (summary && summary.args) {
        var argList = summary.args;
        for (var i = 0, lenI = argList.length; i < lenI; i++) {
          var arg = argList[i];
          if ((arg.type) && (arg.type == 'HYPERLINK') &&
            (arg.key) && (arg.key == 'LINK') &&
              (arg.value)) {
                columnsInfo[part.localizedRuleName] = [arg.value];
              }
        }
      }
      if (part.urlBlocks) {
        var blocks = part.urlBlocks;
        var urls = [];
        for (var j = 0, lenJ = blocks.length; j < lenJ; j++) {
          var block = blocks[j];
          if (block.header) {
            var header = block.header;
            if (header.args) {
              var args = header.args;
              for (var k = 0, lenK = args.length; k < lenK; k++) {
                var argument = args[k];
                if ((argument.type) &&
                    (argument.type == 'HYPERLINK') &&
                  (argument.key) &&
                    (argument.key == 'LINK') &&
                      (argument.value)) {
                        urls.push(argument.value);
                      }
              }
            }
          }
        }
        if (urls.length > 0) {
          columnsInfo[part.localizedRuleName] = urls;
        }
      }
    }
  }
  return columnsInfo;
}

/**
* Extracts a suitable error message to display for a failed URL. The error
* could be passed in in the nested PageSpeed error format, or there could have
* been a more fundamental error in the fetching of the URL. Extract the
* relevant message in each case.
*
* @param {String} errorMessage The error string.
* @return {string} A formatted error message.
*/
function formatErrorMessage(errorMessage) {
  var formattedMessage = null;
  if (!errorMessage) {
    formattedMessage = 'Unknown error message';
  } else {
    try {
      var parsedError = JSON.parse(errorMessage);
      // This is the nested structure expected from PageSpeed
      if (parsedError.error && parsedError.error.errors) {
        var firstError = parsedError.error.errors[0];
        formattedMessage = firstError.message;
      } else if (parsedError.message) {
        formattedMessage = parsedError.message;
      } else {
        formattedMessage = errorMessage.toString();
      }
    } catch (e) {
      formattedMessage = errorMessage.toString();
    }
  }
  return formattedMessage;
}

/**
* Calls the PageSpeed API for a single URL, and attempts to parse the resulting
* JSON. If successful, produces an object for the metrics returned, and an
* object detailing the headings and help URLs seen.
*
* @param {String} url The URL to run PageSpeed for.
* @return {Object} An object with pageSpeed metrics, column-heading info
*     and error properties.
*/
function getPageSpeedResultForSingleUrl(url) {
  var parsedResponse = null;
  var errorMessage = null;
  var retries = 0;
  
  while ((!parsedResponse || parsedResponse.responseCode !== 200) &&
         retries < MAX_RETRIES) {
    errorMessage = null;
    var fetchResult = checkUrl(url);
    if (fetchResult.responseText) {
      try {
        parsedResponse = JSON.parse(fetchResult.responseText);
        break;
      } catch (e) {
        errorMessage = formatErrorMessage(e);
      }
    } else {
      errorMessage = formatErrorMessage(fetchResult.error);
    }
    retries++;
    Utilities.sleep(1000 * Math.pow(2, retries));
  }
  if (!errorMessage) {
    var columnsInfo = extractColumnsInfo(parsedResponse);
    var urlScores = extractResultRow(parsedResponse);
  }
  return {
    pageSpeedInfo: urlScores,
    columnsInfo: columnsInfo,
    error: errorMessage
  };
}

function UrlStore(opt_manualUrls) {
  this.manualUrls = opt_manualUrls || [];
  this.paths = {};
  this.re = /^(https?:\/\/[^\/]+)([^?#]*)(.*)$/;
}

UrlStore.prototype.addUrl = function(url) {
  if (!url || this.manualUrls.indexOf(url) > -1) {
    return;
  }
  var matches = this.re.exec(url);
  if (matches) {
    var host = matches[1];
    var path = matches[2];
    var param = matches[3];
    if (!this.paths[host]) {
      this.paths[host] = {};
    }
    var hostObj = this.paths[host];
    if (!path) {
      path = '/';
    }
    if (!hostObj[path]) {
      hostObj[path] = {};
    }
    var pathObj = hostObj[path];
    pathObj[url] = url;
  }
};

/**
* Adds multiple URLs to the UrlStore.
*
* @param {String[]} urls The URLs to add.
*/
UrlStore.prototype.addUrls = function(urls) {
  for (var i = 0; i < urls.length; i++) {
    this.addUrl(urls[i]);
  }
};

/**
* Creates and returns an iterator that tries to iterate over all available
* URLs return them in an order to maximise the difference between them.
*
* @return {UrlStoreIterator} The new iterator object.
*/
UrlStore.prototype.__iterator__ = function() {
  return new UrlStoreIterator(this.paths, this.manualUrls);
};

var UrlStoreIterator = (function() {
  function UrlStoreIterator(paths, manualUrls) {
    this.manualUrls = manualUrls.slice();
    this.urls = objectToArray_(paths);
  }
  UrlStoreIterator.prototype.next = function() {
    if (this.manualUrls.length) {
      return this.manualUrls.shift();
    }
    if (this.urls.length) {
      return pick_(this.urls);
    } else {
      throw StopIteration;
    }
  };
  function rotate_(a) {
    if (a.length < 2) {
      return a;
    } else {
      var e = a.pop();
      a.unshift(e);
    }
  }
  function pick_(a) {
    if (typeof a[0] === 'string') {
      return a.shift();
    } else {
      var element = pick_(a[0]);
      if (!a[0].length) {
        a.shift();
      } else {
        rotate_(a);
      }
      return element;
    }
  }
  
  function objectToArray_(obj) {
    if (typeof obj !== 'object') {
      return obj;
    }
    
    var a = [];
    for (var k in obj) {
      a.push(objectToArray_(obj[k]));
    }
    return a;
  }
  return UrlStoreIterator;
})();


/**
* Runs the PageSpeed fetch.
*
* @param {String} url
* @return {object} An object containing either the successful response from the
*     server, or an error message.
*/
function checkUrl(url) {
  if(url.indexOf('http') < 0) {
    url = 'https://' + url; 
  }
  
  var result = null;
  var error = null;
  var fullUrl = PAGESPEED_URL + 'key=' + API_KEY + '&url=' + encodeURI(url) +
    '&prettyprint=false&strategy=mobile';
  var params = {muteHttpExceptions: true};
  try {
    var pageSpeedResponse = UrlFetchApp.fetch(fullUrl, params);
    if (pageSpeedResponse.getResponseCode() === 200) {
      result = JSON.parse(pageSpeedResponse.getContentText());
    } else {
      error = pageSpeedResponse.getContentText();
    }
  } catch (e) {
    error = e.message;
  }
  
  if(result && result.ruleGroups) {
    return [[result.ruleGroups['SPEED']['score'], result.ruleGroups['USABILITY']['score']]] ;
  }
  
  return '';
}