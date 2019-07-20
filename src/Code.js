var SCRIPT_PROPERTY_KEYS = {
  OUTPUT_SHEET_ID: 'output_sheet_id',
  NOTIFICATION_MESSAGE: 'notification_message',
  CHATWORK_API_TOKEN: 'chatwork_api_token',
  CHATWORK_ROOM_ID: 'chatwork_room_id'
};

var sheet = (function () {
  var sheets = SpreadsheetApp.getActive().getSheets();
  var sheetId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROPERTY_KEYS.OUTPUT_SHEET_ID);
  if (sheetId) {
    return sheets.filter(function (sheet) {
      return sheet.getId() === sheetId;
    })[0];
  }
  return sheets[0];
}());

var notificationSettings = (function () {
  var scriptProperties = PropertiesService.getScriptProperties();
  var apiToken = scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.CHATWORK_API_TOKEN);
  var roomId = scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.CHATWORK_ROOM_ID);
  var message = scriptProperties.getProperty(SCRIPT_PROPERTY_KEYS.NOTIFICATION_MESSAGE);
  
  if (apiToken == null || roomId == null || message == null) {
    return null;
  }
  
  return {
    apiToken: apiToken,
    roomId: roomId,
    message: message
  };
}());

var searchMessages = function (query, start, max) {
  query = query || '';
  start = start || 0;
  max = max || 100;

  return GmailApp.search(query, start, max).reverse().map(function (gmailThread) {
    return gmailThread.getMessages()[0];
  });
};

var notify = function () {
  if (notificationSettings == null) {
    Logger.log('[Notice] Skip notification due to insufficient settings');
    return;
  }

  ChatWorkClient.factory({
    token: notificationSettings.apiToken
  }).sendMessage({
    room_id: notificationSettings.roomId,
    body: notificationSettings.message
  });
};

function execute() {
  if (sheet == null) throw new Error('[Error] sheet not found');

  var query = [
    'in:inbox',
    'is:unread'
//    'label:example'
  ].join(' ');
  var gmailMessages = searchMessages(query);
  Logger.log('Messages: ' + gmailMessages.length);

  var isDirty = false;

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

    isDirty = true;
    Logger.log('ok');
  }

  if (isDirty) {
    notify();
  }
  
  Logger.log('done');
}
