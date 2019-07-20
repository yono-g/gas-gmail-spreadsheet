var SCRIPT_PROPERTY_KEY_OUTPUT_SHEET_ID = 'output_sheet_id';

var sheet = (function () {
  var sheets = SpreadsheetApp.getActive().getSheets();
  var sheetId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_KEY_OUTPUT_SHEET_ID);
  if (sheetId) {
    return sheets.filter(function (sheet) {
      return sheet.getId() === sheetId;
    })[0];
  }
  return sheets[0];
}());

var searchMessages = function (query, start, max) {
  query = query || '';
  start = start || 0;
  max = max || 100;

  return GmailApp.search(query, start, max).reverse().map(function (gmailThread) {
    return gmailThread.getMessages()[0];
  });
};

var notify = function () {};

function execute() {
  if (sheet == null) throw new Error('Error: sheet not found');

  var query = [
    'in:inbox',
    'is:unread'
//    'label:example'
  ].join(' ');

  var notifyNeeded = false;

  var gmailMessages = searchMessages(query);
  Logger.log('Messages: ' + gmailMessages.length);
  for (var i=0; i<gmailMessages.length; i++) {
    var gmailMessage = gmailMessages[i];
    Logger.log('Message: ' + gmailMessage.getId());

    var date = gmailMessage.getDate();
    var from = gmailMessage.getFrom();
    var plainBody = gmailMessage.getPlainBody();
    
    sheet.appendRow([
      date,
      from,
      plainBody
    ]);

    gmailMessage.markRead();
    notifyNeeded = true;
    Logger.log('ok');
  }

  if (notifyNeeded) {
    notify();
  }
}
