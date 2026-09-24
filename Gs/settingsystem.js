var settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
var idfolder = settingSheet.getRange('B1').getDisplayValue();
var sigFolder = settingSheet.getRange('B2').getDisplayValue();
var imageFolder = settingSheet.getRange('B3').getDisplayValue();
var sheetDataA = settingSheet.getRange('B4').getDisplayValue();
var sheetDataB = settingSheet.getRange('B5').getDisplayValue();
var sheetDataC = settingSheet.getRange('B6').getDisplayValue();
var sheetDataDEF = settingSheet.getRange('B7').getDisplayValue();
var sheetDataG = settingSheet.getRange('B8').getDisplayValue();
var sheetDataH = settingSheet.getRange('B9').getDisplayValue();
var sheetDataSA = settingSheet.getRange('B10').getDisplayValue();
var sheetDataSB = settingSheet.getRange('B11').getDisplayValue();
var sheetDataSearch = settingSheet.getRange('B12').getDisplayValue();
var sheetDataSet = settingSheet.getRange('B13').getDisplayValue();
var logoUrl = settingSheet.getRange('B14').getDisplayValue();
var nameSystem = settingSheet.getRange('B15').getDisplayValue();

const getSet = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  var data = ss.getRange("B1:B").getDisplayValues();
  return data;
};

const settingGS = (data) => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Setting');
  var valuesToSet = [];
  for (let i = 1; i <= 16; i++) {
    valuesToSet.push([data[`set${i}`]]);
  }
  var range = sheet.getRange(1, 2, valuesToSet.length, 1);
  range.setValues(valuesToSet);
}

const selectDataFromSheet = (sheetName) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(sheetName);
  var getLastRow = sheet.getLastRow();
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
  var data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
  return data.flat().filter(Boolean);
}

const saveTodos = (data) => {
  var sheet = SpreadsheetApp.openById(sheetDataSet).getSheetByName(data.sheetName);
  sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).clearContent();
  data.todos.forEach((todo, index) => {
    sheet.getRange(index + 2, 2).setValue(todo);
  });
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