/*var SHEET_ID = '1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM';
var SHEET_NAME = 'APP 推播表';
var HISTORY_SHEET_NAME = 'APP推播歷史';
var PROJECT_ID = 'fcm-playbeautyshop';
var credentials = getCredentials();

function sendNotificationFromSheet() {
  // 暫停所有現有的觸發器
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
      var payload;
      if (to === 'all') {
        payload = {
          'message': {
            'topic': 'all',
            'notification': {
              'title': title,
              'body': body,
              'image': imageUrl
            },
            'apns': {
              'payload': {
                'aps': {
                  'alert': {
                    'title': title,
                    'body': body
                  },
                  'sound': 'default',
                  'badge': 1
                },
                'url': action
              },
              'headers': {
                'apns-push-type': 'alert',
                'apns-priority': '10',
                'apns-topic': 'com.playbeautyshop.playbeautyshopapp'
              }
            },
            'android': {
              'priority': priority,
              'notification': {
                'title': title,
                'body': body,
                'image': imageUrl,
              },
              'data': {
                'url': action
              }
            }
          }
        };
      } else {
        payload = {
          'message': {
            'token': to,
            'notification': {
              'title': title,
              'body': body,
              'image': imageUrl
            },
            'apns': {
              'payload': {
                'aps': {
                  'alert': {
                    'title': title,
                    'body': body
                  },
                  'sound': 'default',
                  'badge': 1
                },
                'url': action
              },
              'headers': {
                'apns-push-type': 'alert',
                'apns-priority': '10',
                'apns-topic': 'com.playbeautyshop.playbeautyshopapp'
              }
            },
            'android': {
              'priority': priority,
              'notification': {
                'title': title,
                'body': body,
                'image': imageUrl,
              },
              'data': {
                'url': action
              }
            }
          }
        };
      }

      httpOptions.payload = JSON.stringify(payload);
      var url = 'https://fcm.googleapis.com/v1/projects/' + PROJECT_ID + '/messages:send';
      UrlFetchApp.fetch(url, httpOptions);
      sheet.getRange(i + 1, 9).setValue("done");

      // 將已執行的任務移動到 "APP推播歷史" 工作表,並將 D 欄的時間更新為當前時間
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
}*/