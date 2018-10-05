var EMAIL = 'manoj.gohil@eagerhealth.com,kieran.sheedy@eagerhealth.com';

function main() {
  MccApp.accounts().withCondition('Name = "C247 - Chiswick"').executeInParallel('run');
}

function run() {
  
  var hour = parseInt(Utilities.formatDate(new Date(), AdWordsApp.currentAccount().getTimeZone(), 'HH'),10);
  hour++;
  if(hour == 24) { hour = 0; }
  if(hour >= 8 && hour < 18) { return; }
  
  if(hour > 12) {
    hour = (hour-12) + ' PM'; 
  } else {
    if(hour == 0) { hour = 12; }
    hour = hour + ' AM'; 
  }
  
  var spend = AdWordsApp.currentAccount().getStatsFor('TODAY').getCost();
  var MSG = 'Hi,\n\n' + hour + ': Spends for the Account - £'+spend + '\n\nThanks';
  
  MailApp.sendEmail(EMAIL, 'Care24Seven: Spends Report (Today)', MSG);
}