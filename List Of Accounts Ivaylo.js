var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1mR8VrUpqit0Tf6McDGAu071Xsdotfh4M1U1vX-tLUMI/edit#gid=0';
var SHEET_NAME = 'Sheet1';

function main() {
  
  var accountIterator = MccApp.accounts().get();

  while (accountIterator.hasNext()) {
    
    var account = accountIterator.next();
    
    var accountCid = account.getCustomerId();
    var accountName = account.getName() ? account.getName() : '--';
    var accountCurrencyCode = account.getCurrencyCode();
    var accountTimezone = account.getTimeZone();

	var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
	var sheet = ss.getSheetByName(SHEET_NAME);

	// Appends a new row with 3 columns to the bottom of the
	// spreadsheet containing the values in the array.
	sheet.appendRow([accountCid, accountName, accountCurrencyCode, accountTimezone]);

  }
}

