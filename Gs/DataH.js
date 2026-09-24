const getDataH = () => {
  const sheetH = SpreadsheetApp.openById(sheetDataH);
  const sheet = sheetH.getSheetByName("DataH"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  //Logger.log(data)
  return data;
}

function generateIDDataH(currentIDDataH) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataH.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const addDataH = (obj) => {
  const sheetH = SpreadsheetApp.openById(sheetDataH).getSheetByName('DataH')
  const lastRowID = sheetH.getLastRow();
  const codeIDDataH = generateIDDataH(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => {
    return fileData.length > 0 ? documentFolder.createFile(fileData.setName(namePrefix + codeIDDataH)).getUrl() : "";
  };
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataH5, "ประชาสัมพันธ์(ไฟล์หลัก)")|| "";
  const d = new Date(); 
  const curdate = d.getDate(); 
  const monthcur = d.getMonth(); 
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;

  const rowData = ["'"+codeIDDataH, obj.dataDataHInput1, obj.dataDataHInput2, obj.dataDataHInput3, newdateThai, ucfileA1];
  sheetH.appendRow(rowData);

  const tinyurlPDF = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + codeIDDataH +
              '\n📝 เรื่อง: ' + obj.dataDataHInput1 +
              '\n📝 รายละเอียด: ' + obj.dataDataHInput2 +
              '\n💻 ผู้โพสต์: ' + obj.dataDataHInput3 +
              '\n📆 วันที่โพสต์: ' + newdateThai +
              '\n📁 ไฟล์ PDF: ' + tinyurlPDF

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetH.getRange("A2:F" + sheetH.getLastRow()).getValues();
}

const upDateDataH = (obj) => {
  const sheetH = SpreadsheetApp.openById(sheetDataH).getSheetByName('DataH')
  const data = sheetH.getDataRange().getDisplayValues()
  const iddataH = data.map(r=>r[0])
  const rowIndex = iddataH.indexOf(obj.dataDataHInputKey);
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataH5.length > 0) {
    ucfileA1 = documentFolder.createFile(obj.myfileDataH5.setName("ประชาสัมพันธ์(ไฟล์หลัก)" + obj.dataDataHInputKey)).getUrl();
    const oldfile = sheetH.getRange(rowIndex + 1, 6).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetH.getRange(rowIndex + 1, 6).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetH.getRange(rowIndex + 1, 2).setValue(obj.dataDataHInput1);
  sheetH.getRange(rowIndex + 1, 3).setValue(obj.dataDataHInput2);
  sheetH.getRange(rowIndex + 1, 4).setValue(obj.dataDataHInput3);
  sheetH.getRange(rowIndex + 1, 5).setValue(obj.dataDataHInput4);
  }

  const tinyurlPDF = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + obj.dataDataHInputKey +
              '\n📝 เรื่อง: ' + obj.dataDataHInput1 +
              '\n📝 รายละเอียด: ' + obj.dataDataHInput2 +
              '\n💻 ผู้โพสต์: ' + obj.dataDataHInput3 +
              '\n📆 วันที่โพสต์: ' + obj.dataDataHInput4 +
              '\n📁 ไฟล์ PDF: ' + tinyurlPDF

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetH.getRange("A2:G" + sheetH.getLastRow()).getValues();
}

const delRecDataH = (record) =>  {
  const sheetH = SpreadsheetApp.openById(sheetDataH).getSheetByName('DataH')
  const data = sheetH.getDataRange().getDisplayValues()
  const iddataH = data.map(r=>r[0])
  const rowIndex = iddataH.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetH.getRange(rowIndex + 1, 6).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetH.deleteRow(rowIndex + 1);
  }
}