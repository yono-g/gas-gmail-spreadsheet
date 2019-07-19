var SCRIPT_PROPERTY_KEY_DEST_SHEET_ID = 'dest_sheet_id';

function myFunction() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty(SCRIPT_PROPERTY_KEY_DEST_SHEET_ID);

  var sheets = SpreadsheetApp.getActive().getSheets();
  var sheet = sheetId ? sheets.filter(function (s) { return s.getId() === sheetId; }) : sheets[0];
  if (!sheet) {
    throw new Error('sheet not found');
  }

  var query = [
    'in:inbox',
    'is:unread',
    'label:example'
  ].join(' ');

  Logger.log('Query: "' + query + '"');

  var isDirty = false;

  var gmailThreads = GmailApp.search(query);  // TODO: Consider start, max
  Logger.log('Results: ' + gmailThreads.length);

  for (var i=0; i<gmailThreads.length; i++) {
    Logger.log('Thread ID: ' + gmailThreads[i].getId());

    var gmailMessage = gmailThreads[i].getMessages()[0];
    if (gmailMessage == null) {
      Logger.log('\tno first message.');
      continue;
    }
    Logger.log('\tMessage ID: ' + gmailMessage.getId());

    var data = [
      gmailMessage.getDate(),
      gmailMessage.getFrom(),
      gmailMessage.getPlainBody()
    ];
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow+1, 1, 1, data.length).setValues([data]);

    gmailThreads[i].markRead();

    isDirty = true;
  }

  if (isDirty) {
    // TODO: notify changes to a chat service or something else.
    Logger.log('notify feature is not implemented yet');
  }
}
