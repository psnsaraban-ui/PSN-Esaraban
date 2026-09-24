// let sheetG = SpreadsheetApp.openById(sheetDataG).getSheetByName('DataG')
// let data = sheetG.getDataRange().getDisplayValues()
// let iddataG = data.map(r=>r[0])

const getDataG = () => {
  const sheetG = SpreadsheetApp.openById(sheetDataG);
  const sheet = sheetG.getSheetByName("DataG"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  //Logger.log(data)
  return data;
}

function generateIDDataG(currentIDDataG) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataG.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const addDataG = (obj) => {
  const sheetG = SpreadsheetApp.openById(sheetDataG).getSheetByName('DataG');
  const lastRowID = sheetG.getLastRow();
  const codeIDDataG = generateIDDataG(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => createPdfFileFromBase64(documentFolder, fileData, namePrefix + codeIDDataG);
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataG5, "ประกาศ(ไฟล์หลัก)")|| "";
  const d = new Date(); 
  const curdate = d.getDate(); 
  const monthcur = d.getMonth(); 
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;

  const rowData = ["'" + codeIDDataG, "รอประกาศ", obj.dataDataGInput2, newdateThai,  obj.dataDataGInput4, ucfileA1];
  sheetG.appendRow(rowData);

  const tinyurlPDF = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + codeIDDataG +
              '\n📍 ลงวันที่: ' + newdateThai +
              '\n📝 เรื่อง: ' + obj.dataDataGInput2 +
              '\n💻 หน่วยงาน: ' + obj.dataDataGInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataGInput4 +
              '\n📁 ไฟล์ PDF: ' + tinyurlPDF

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetG.getRange("A2:F" + sheetG.getLastRow()).getValues();
}

const upDateDataG = (obj) => {
  const sheetG = SpreadsheetApp.openById(sheetDataG).getSheetByName('DataG');
  const data = sheetG.getDataRange().getDisplayValues();
  const iddataG = data.map(r => r[0]);
  const rowIndex = iddataG.indexOf(obj.dataDataGInputKey);
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataG5) {
    ucfileA1 = createPdfFileFromBase64(documentFolder, obj.myfileDataG5, "ประกาศ(ไฟล์หลัก)" + obj.dataDataGInputKey);
    const oldfile = sheetG.getRange(rowIndex + 1, 6).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetG.getRange(rowIndex + 1, 6).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetG.getRange(rowIndex + 1, 2).setValue(obj.dataDataGInput1);
  sheetG.getRange(rowIndex + 1, 3).setValue(obj.dataDataGInput2);
  sheetG.getRange(rowIndex + 1, 4).setValue(obj.dataDataGInput3);
  sheetG.getRange(rowIndex + 1, 5).setValue(obj.dataDataGInput4);
  }

  const tinyurlPDF = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + obj.dataDataGInputKey +
              '\n📍 ลงวันที่: ' + obj.dataDataGInput1 +
              '\n📝 เรื่อง: ' + obj.dataDataGInput2 +
              '\n💻 หน่วยงาน: ' + obj.dataDataGInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataGInput4 +
              '\n📁 ไฟล์ PDF: ' + tinyurlPDF

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetG.getRange("A2:F" + sheetG.getLastRow()).getValues();
}

const delRecDataG = (record) =>  {
  const sheetG = SpreadsheetApp.openById(sheetDataG).getSheetByName('DataG');
  const data = sheetG.getDataRange().getDisplayValues();
  const iddataG = data.map(r => r[0]);
  const rowIndex = iddataG.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetG.getRange(rowIndex + 1, 6).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetG.deleteRow(rowIndex + 1);
  }
}