var settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
var idfolder = settingSheet.getRange('B1').getDisplayValue();
var sigFolder = settingSheet.getRange('B2').getDisplayValue();
var imageFolder = settingSheet.getRange('B3').getDisplayValue();
var dataSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
var sheetDataA = dataSpreadsheetId;
var sheetDataB = dataSpreadsheetId;
var sheetDataC = dataSpreadsheetId;
var sheetDataDEF = dataSpreadsheetId;
var sheetDataG = dataSpreadsheetId;
var sheetDataH = dataSpreadsheetId;
var sheetDataSA = dataSpreadsheetId;
var sheetDataSB = dataSpreadsheetId;
var sheetDataSearch = dataSpreadsheetId;
var sheetDataSet = dataSpreadsheetId;
var logoUrl = settingSheet.getRange('B14').getDisplayValue();
var nameSystem = settingSheet.getRange('B15').getDisplayValue();

function migrateDataToSingleSpreadsheet() {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const masterId = master.getId();
  const settings = master.getSheetByName('Setting');
  const sourceIds = settings.getRange('B4:B13').getDisplayValues().flat();
  const sourceTabs = [
    ['DataA'],
    ['DataB'],
    ['DataC'],
    ['DataD', 'DataE', 'DataF'],
    ['DataG'],
    ['DataH'],
    ['DataSA'],
    ['DataSB'],
    ['DataSearch'],
    ['Agency', 'Position', 'Department', 'Objective', 'ClassSpeed', 'ClassSecret', 'Response']
  ];
  const migrationPlan = [];
  const hasSameValues = (leftSheet, rightSheet) =>
    JSON.stringify(leftSheet.getDataRange().getValues()) === JSON.stringify(rightSheet.getDataRange().getValues());

  sourceTabs.forEach((tabNames, index) => {
    const sourceId = sourceIds[index] || masterId;
    const source = SpreadsheetApp.openById(sourceId);
    tabNames.forEach(name => {
      const sourceSheet = source.getSheetByName(name);
      if (!sourceSheet) {
        throw new Error('Source tab not found: ' + name);
      }
      const targetSheet = master.getSheetByName(name);
      const alreadyCopied = sourceId !== masterId && targetSheet && targetSheet.getLastRow() > 0;
      if (alreadyCopied && !hasSameValues(sourceSheet, targetSheet)) {
        throw new Error('Target tab already contains different data: ' + name);
      }
      migrationPlan.push({ sourceId, sourceSheet, targetSheet, name, alreadyCopied });
    });
  });

  const copiedTabs = [];
  migrationPlan.forEach(item => {
    if (item.sourceId === masterId || item.alreadyCopied) return;
    if (item.targetSheet) master.deleteSheet(item.targetSheet);
    item.sourceSheet.copyTo(master).setName(item.name);
    copiedTabs.push(item.name);
  });

  settings.getRange('B4:B13').setValues(Array.from({ length: 10 }, () => [masterId]));
  return { spreadsheetId: masterId, copiedTabs };
}

const getSet = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ss = spreadsheet.getSheetByName('Setting');
  const data = ss.getRange('B1:B16').getDisplayValues();
  for (let index = 3; index <= 12; index++) {
    data[index] = [spreadsheet.getId()];
  }
  return data;
};

const settingGS = (data) => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Setting');
  var valuesToSet = [];
  for (let i = 1; i <= 16; i++) {
    valuesToSet.push([i >= 4 && i <= 13 ? ss.getId() : data[`set${i}`]]);
  }
  var range = sheet.getRange(1, 2, valuesToSet.length, 1);
  range.setValues(valuesToSet);
}

const selectDataFromSheet = (sheetName) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(sheetName);
  var getLastRow = sheet.getLastRow();
  if (getLastRow < 2) return [];
  var data = sheet.getRange(2, 2, getLastRow - 1, 1).getValues().flat();
  return data;
}

const selectAgency = () => selectDataFromSheet("Agency");
const selectPosition = () => selectDataFromSheet("Position");
const selectDepartment = () => selectDataFromSheet("Department");
const selectObjectivet = () => selectDataFromSheet("Objective");
const selectClassSpeed = () => selectDataFromSheet("ClassSpeed");
const selectClassSecret = () => selectDataFromSheet("ClassSecret");
const selectResponse = () => selectDataFromSheet("Response");

const getTodos = (sheetName) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(sheetName);
  if (sheet.getLastRow() < 2) return [];
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  return data.flat().filter(Boolean);
}

const saveTodos = (data) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(data.sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 2, lastRow - 1, 1).clearContent();
  if (data.todos.length > 0) {
    sheet.getRange(2, 2, data.todos.length, 1).setValues(data.todos.map(todo => [todo]));
  }
}

const selectUsersToken = () => {
  var sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data = sheetUsers.getRange('A2:H' + sheetUsers.getLastRow()).getValues();
  var filteredData = data.filter(function(row) {
    return row[0] !== '';
  });
  var departmentsOrder = selectDepartment();
  var result = filteredData.map(function(row) {
    return {
      sendID: row[0],
      sendFullname: row[3],
      sendDepartment: row[4]
    };
  });
  
  result.sort(function(a, b) {
    var indexA = departmentsOrder.indexOf(a.sendDepartment);
    var indexB = departmentsOrder.indexOf(b.sendDepartment);
    return indexA - indexB;
  });
  
  //Logger.log(result)
  return result;
}

const selectUsersAgency = () => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName("Agency");
  var data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
  var filteredData = data.filter(function(row) {
    return row[0] !== '';
  });
  var result = filteredData.map(function(row) {
    return {
      agencyID: row[0],
      agencyName: row[1]
    };
  });
  
  //Logger.log(result);
  return result;
}

function shortenURL(longURL) {
  var apiUrl = "http://tinyurl.com/api-create.php?url=" + encodeURIComponent(longURL);
  var response = UrlFetchApp.fetch(apiUrl);
  return response.getContentText();
}

function getNotificationSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notification');
  const data = sheet.getDataRange().getValues();
  const settings = [];
  
  for (let i = 1; i < data.length; i++) {
    settings.push({
      id: String(data[i][0]),
      type: data[i][1],
      token: data[i][2],
      chatId: data[i][3],
      status: data[i][4]
    });
  }

  return settings.filter(setting => setting.status === true);
}

function sendNotify(msg, setting, imageUrls = []) {
  const validImageUrls = imageUrls.filter(url => url && url.trim() !== "");

  if (setting.id === "1") {  // MessagingAPI
    if (validImageUrls.length > 0) {
      let payloadJson = {
        "to": setting.chatId,
        "messages": [
          {
            "type": "text",
            "text": msg
          },
          {
            "type": "image",
            "originalContentUrl": validImageUrls[0],
            "previewImageUrl": validImageUrls[0]
          }
        ]
      };
      let options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payloadJson),
        "headers": {
          "Authorization": "Bearer " + setting.token
        }
      };
      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
    } else {
      let payloadJson = {
        "to": setting.chatId,
        "messages": [
          {
            "type": "text",
            "text": msg
          }
        ]
      };
      let options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payloadJson),
        "headers": {
          "Authorization": "Bearer " + setting.token
        }
      };
      UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
    }
  } 
  else if (setting.id === "2") { // Telegram
    if (validImageUrls.length > 0) {
      validImageUrls.forEach((imageUrl, index) => {
        let payloadJson = {
          "chat_id": setting.chatId,
          "photo": imageUrl,
          "caption": index === 0 ? msg : "",
          "parse_mode": "Markdown"
        };
        let options = {
          "method": "post",
          "contentType": "application/json",
          "payload": JSON.stringify(payloadJson)
        };
        UrlFetchApp.fetch("https://api.telegram.org/bot" + setting.token + "/sendPhoto", options);
      });
    } else {
      let payloadJson = {
        "chat_id": setting.chatId,
        "text": msg,
        "parse_mode": "Markdown"
      };
      let options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payloadJson)
      };
      UrlFetchApp.fetch("https://api.telegram.org/bot" + setting.token + "/sendMessage", options);
    }
  } 
  else if (setting.id === "3") { // Discord
    let payloadJson = {
      "content": msg
    };
    if (validImageUrls.length > 0) {
      payloadJson.embeds = validImageUrls.map(url => ({
        "image": {
          "url": url
        }
      }));
    }
    let options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payloadJson)
    };
    UrlFetchApp.fetch(setting.token, options);
  }
}

function testTelegramBot() {
  const telegramSetting = getNotificationSettings().find(setting => setting.id === "2");
  if (!telegramSetting) {
    throw new Error('ไม่พบการตั้งค่า Telegram ที่เปิดใช้งานในชีต Notification');
  }

  sendNotify('ทดสอบการส่งข้อความจาก I-OFFICE สำเร็จ', telegramSetting);
  return 'ส่งข้อความทดสอบไปยัง Telegram แล้ว';
}

// function sendNotify(msg, tokens, imgUrl) {
//     let payloadJson = {
//         "message": msg
//     };
//     if (imgUrl) {
//         payloadJson.imageThumbnail = imgUrl;
//         payloadJson.imageFullsize = imgUrl;
//     }
//     let options = {
//         "method": "post",
//         "payload": payloadJson,
//         "headers": {
//             "Authorization": "Bearer " + tokens
//         }
//     };
//     UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }