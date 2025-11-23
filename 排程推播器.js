var SHEET_ID = '1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM';
var SHEET_NAME = 'APP 推播表';
var HISTORY_SHEET_NAME = 'APP推播歷史';
var PROJECT_ID = 'fcm-playbeautyshop';
var credentials = getCredentials();

// === 新增 buildPayload ===
function buildPayload({to, title, body, imageUrl, action, priority, notifId}) {
  const apnsHeaders = {
    'apns-push-type': 'alert',
    'apns-priority': '10',
    'apns-topic': 'com.playbeautyshop.playbeautyshopapp',
    // 🎯 每則通知都不同，避免覆蓋
    'apns-collapse-id': notifId
  };

  const apnsPayload = {
    'aps': {
      'alert': { 'title': title, 'body': body },
      'sound': 'default',
      'badge': 1,
      'thread-id': 'promo',        // 群組用
      'mutable-content': 1         // 支援圖片/富媒體
    },
    'url': action || ''
  };

  const base = {
    'notification': { 'title': title, 'body': body, 'image': imageUrl },
    'apns': { 'headers': apnsHeaders, 'payload': apnsPayload },
    'android': {
      'priority': priority || 'HIGH',
      'notification': { 'title': title, 'body': body, 'image': imageUrl },
      'data': { 'url': action || '' }
    }
  };

  return {
    'message': (to === 'all')
      ? { ...base, 'topic': 'all' }
      : { ...base, 'token': to }
  };
}

// === 主流程 ===
function sendNotificationFromSheet() {
  // 暫停所有現有的觸發器 (保留原本程式邏輯)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  var historySheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var httpOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + getAccessToken()
    }
  };
  var now = new Date();

  for (var i = data.length - 1; i > 0; i--) {
    var title = data[i][0];
    var body = data[i][1];
    var to = data[i][2];
    var scheduledTimeStr = data[i][3];
    var imageUrl = data[i][4];
    var action = data[i][5];
    var priority = data[i][6];
    var task = data[i][7];
    var sendNotification = false;

    if (task === "push") {
      if (scheduledTimeStr === 'now') {
        sendNotification = true;
      } else {
        var scheduledDate = new Date(scheduledTimeStr);
        if (now >= scheduledDate) {
          sendNotification = true;
        }
      }
    }

    if (sendNotification) {
      // === 生成唯一通知 ID ===
      var notifId = 'pb_' + i + '_' + now.getTime();

      // === 使用 buildPayload ===
      var payload = buildPayload({
        to, title, body, imageUrl, action, priority, notifId
      });

      httpOptions.payload = JSON.stringify(payload);
      var url = 'https://fcm.googleapis.com/v1/projects/' + PROJECT_ID + '/messages:send';
      UrlFetchApp.fetch(url, httpOptions);
      sheet.getRange(i + 1, 9).setValue("done");

      // 移動到歷史 sheet 並更新時間
      var rowData = data[i];
      rowData[3] = new Date();
      historySheet.appendRow(rowData);
      sheet.deleteRow(i + 1);
    }
  }

  // 重新創建定時觸發器
  ScriptApp.newTrigger('sendNotificationFromSheet')
    .timeBased()
    .everyMinutes(10)
    .create();
}

function getAccessToken() {
  var serviceAccount = JSON.parse(JSON.stringify(credentials));
  var scope = 'https://www.googleapis.com/auth/firebase.messaging';
  return OAuth2.createService('FCM')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setPrivateKey(serviceAccount.private_key)
    .setIssuer(serviceAccount.client_email)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(scope)
    .getAccessToken();
}
