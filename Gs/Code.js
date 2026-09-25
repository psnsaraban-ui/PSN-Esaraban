function doGet(){
  return HtmlService.createTemplateFromFile('AppsScript').evaluate()
  .setTitle(nameSystem)
  .setFaviconUrl(logoUrl)
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const functionName = String(request.functionName || '');
    if (!/^[A-Za-z_$][\w$]*$/.test(functionName)) {
      throw new Error('Invalid function name');
    }

    const targetFunction = eval(functionName);
    if (typeof targetFunction !== 'function') {
      throw new Error('Function not found: ' + functionName);
    }

    const result = targetFunction.apply(null, Array.isArray(request.args) ? request.args : []);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result: result === undefined ? null : result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

const getURL = () => {
  return ScriptApp.getService().getUrl();
}

const include = (filename) => {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

const createPdfFileFromBase64 = (folder, base64Data, filename) => {
  if (!base64Data) return "";
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, MimeType.PDF, filename);
  return folder.createFile(blob).getUrl();
}

const getDataSearch = () => {
  const sheet = SpreadsheetApp.openById(sheetDataSearch);
  const sheetset = sheet.getSheetByName("DataSearch"); 
  const data = sheetset.getDataRange().getDisplayValues().slice(1);
  //Logger.log(data)
  return data;
}

const getsetMenuItems = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  return data;
}

const getMenuItems = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const menuItems = {};

  for (let i = 1; i < data.length; i++) {
    const item = data[i][1];
    menuItems[item] = {};
    for (let j = 2; j < headers.length; j++) { 
      const cellValue = String(data[i][j]).toUpperCase() || "FALSE"; 
      menuItems[item][headers[j]] = cellValue === "TRUE";
    }
  }
  return menuItems;
};

const updateMenuStatus = (menuItem, role, isChecked) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UsersMenu");
  const data = sheet.getDataRange().getValues();
  const index = data.findIndex(row => row[1] === menuItem);
  if (index !== -1) {
    const roleColumn = role === 'SuperAdmin' ? 3 : role === 'Admin' ? 4 : role === 'SuperUser' ? 5 : 6;
    const range = sheet.getRange(index + 1, roleColumn);
    range.setValue(isChecked ? "TRUE" : "FALSE");
  }
}