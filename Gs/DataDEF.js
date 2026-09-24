// const getDataD = () => {
//   const sheetD = SpreadsheetApp.openById(sheetDataDEF);
//   const sheet = sheetD.getSheetByName("DataD"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataD = (userName, userStatus) => {
  const sheetD = SpreadsheetApp.openById(sheetDataDEF);
  const sheet = sheetD.getSheetByName("DataD");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[5] === userName);
  }
};

// const getDataE = () => {
//   const sheetE = SpreadsheetApp.openById(sheetDataDEF);
//   const sheet = sheetE.getSheetByName("DataE"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataE = (userName, userStatus) => {
  const sheetE = SpreadsheetApp.openById(sheetDataDEF);
  const sheet = sheetE.getSheetByName("DataE"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[5] === userName);
  }
};

// const getDataF = () => {
//   const sheetF = SpreadsheetApp.openById(sheetDataDEF);
//   const sheet = sheetF.getSheetByName("DataF"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataF = (userName, userStatus) => {
  const sheetF = SpreadsheetApp.openById(sheetDataDEF);
  const sheet = sheetF.getSheetByName("DataF"); 
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[5] === userName);
  }
};

function generateIDDataD(currentIDDataD) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataD.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

function generateIDDataE(currentIDDataE) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataE.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

function generateIDDataF(currentIDDataF) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataF.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const addDataD = (obj) => {
  const sheetD = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataD');
  const lastRowID = sheetD.getLastRow();
  const codeIDDataD = generateIDDataD(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => {
    return fileData.length > 0 ? documentFolder.createFile(fileData.setName(namePrefix + codeIDDataD)).getUrl() : "";
  };
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataD6, "หนังสือบันทึกข้อความ(ไฟล์หลัก)")|| "";
  const d = new Date(); 
  const curdate = d.getDate(); 
  const monthcur = d.getMonth(); 
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;
  const rowData = ["'"+codeIDDataD, "รอดำเนินการ", newdateThai, obj.dataDataDInput3, obj.dataDataDInput4, obj.dataDataDInput5, ucfileA1];
  sheetD.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + codeIDDataD +
              '\n📍 ลงวันที่: ' + newdateThai +
              '\n📝 เรื่อง: ' + obj.dataDataDInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataDInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataDInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetD.getRange("A2:G" + sheetD.getLastRow()).getValues();
}

const upDateDataD = (obj) => {
  const sheetD = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataD');
  const data = sheetD.getDataRange().getDisplayValues()
  const iddataD = data.map(r=>r[0])
  const rowIndex = iddataD.indexOf(obj.dataDataDInputKey)
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataD6.length > 0) {
    ucfileA1 = documentFolder.createFile(obj.myfileDataD6.setName("หนังสือบันทึกข้อความ(ไฟล์หลัก)" + obj.dataDataDInputKey)).getUrl();
    const oldfile = sheetD.getRange(rowIndex + 1, 7).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetD.getRange(rowIndex + 1, 7).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetD.getRange(rowIndex + 1, 2).setValue(obj.dataDataDInput1);
  sheetD.getRange(rowIndex + 1, 3).setValue(obj.dataDataDInput2);
  sheetD.getRange(rowIndex + 1, 4).setValue(obj.dataDataDInput3);
  sheetD.getRange(rowIndex + 1, 5).setValue(obj.dataDataDInput4);
  sheetD.getRange(rowIndex + 1, 6).setValue(obj.dataDataDInput5);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + obj.dataDataDInputKey +
              '\n📍 ลงวันที่: ' + obj.dataDataDInput2 +
              '\n📝 เรื่อง: ' + obj.dataDataDInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataDInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataDInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetD.getRange("A2:G" + sheetD.getLastRow()).getValues();
}

const delRecDataD = (record) =>  {
  const sheetD = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataD');
  const data = sheetD.getDataRange().getDisplayValues()
  const iddataD = data.map(r=>r[0])
  var rowIndex = iddataD.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetD.getRange(rowIndex + 1, 7).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetD.deleteRow(rowIndex + 1);
  }
}

const addDataE = (obj) => {
  const sheetE = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataE');
  const lastRowID = sheetE.getLastRow();
  const codeIDDataE = generateIDDataE(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => {
    return fileData.length > 0 ? documentFolder.createFile(fileData.setName(namePrefix + codeIDDataE)).getUrl() : "";
  };
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataE6, "หนังสือคำสั่ง(ไฟล์หลัก)")|| "";
  const d = new Date(); 
  const curdate = d.getDate(); 
  const monthcur = d.getMonth(); 
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;
  const rowData = ["'"+codeIDDataE, "รอดำเนินการ", newdateThai, obj.dataDataEInput3, obj.dataDataEInput4, obj.dataDataEInput5, ucfileA1];
  sheetE.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + codeIDDataE +
              '\n📍 ลงวันที่: ' + newdateThai +
              '\n📝 เรื่อง: ' + obj.dataDataEInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataEInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataEInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetE.getRange("A2:G" + sheetE.getLastRow()).getValues();
}

const upDateDataE = (obj) => {
  const sheetE = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataE');
  const data = sheetE.getDataRange().getDisplayValues()
  const iddataE = data.map(r=>r[0])
  const rowIndex = iddataE.indexOf(obj.dataDataEInputKey)
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataE6.length > 0) {
    ucfileA1 = documentFolder.createFile(obj.myfileDataE6.setName("หนังสือคำสั่ง(ไฟล์หลัก)" + obj.dataDataEInputKey)).getUrl();
    const oldfile = sheetE.getRange(rowIndex + 1, 7).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetE.getRange(rowIndex + 1, 7).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetE.getRange(rowIndex + 1, 2).setValue(obj.dataDataEInput1);
  sheetE.getRange(rowIndex + 1, 3).setValue(obj.dataDataEInput2);
  sheetE.getRange(rowIndex + 1, 4).setValue(obj.dataDataEInput3);
  sheetE.getRange(rowIndex + 1, 5).setValue(obj.dataDataEInput4);
  sheetE.getRange(rowIndex + 1, 6).setValue(obj.dataDataEInput5);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + obj.dataDataEInputKey +
              '\n📍 ลงวันที่: ' + obj.dataDataEInput2 +
              '\n📝 เรื่อง: ' + obj.dataDataEInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataEInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataEInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetE.getRange("A2:G" + sheetE.getLastRow()).getValues();
}

const delRecDataE = (record) =>  {
  const sheetE = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataE');
  const data = sheetE.getDataRange().getDisplayValues()
  const iddataE = data.map(r=>r[0])
  var rowIndex = iddataE.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetE.getRange(rowIndex + 1, 7).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetE.deleteRow(rowIndex + 1);
  }
}

const addDataF = (obj) => {
  const sheetF = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataF');
  const lastRowID = sheetF.getLastRow();
  const codeIDDataF = generateIDDataF(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => {
    return fileData.length > 0 ? documentFolder.createFile(fileData.setName(namePrefix + codeIDDataF)).getUrl() : "";
  };
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataF6, "หนังสือสัญญา(ไฟล์หลัก)")|| "";
  const d = new Date(); 
  const curdate = d.getDate(); 
  const monthcur = d.getMonth(); 
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;
  const rowData = ["'"+codeIDDataF, "รอดำเนินการ", newdateThai, obj.dataDataFInput3, obj.dataDataFInput4, obj.dataDataFInput5, ucfileA1];
  sheetF.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + codeIDDataF +
              '\n📍 ลงวันที่: ' + newdateThai +
              '\n📝 เรื่อง: ' + obj.dataDataFInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataFInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataFInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetF.getRange("A2:G" + sheetF.getLastRow()).getValues();
}

const upDateDataF = (obj) => {
  const sheetF = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataF');
  const data = sheetF.getDataRange().getDisplayValues()
  const iddataF = data.map(r=>r[0])
  const rowIndex = iddataF.indexOf(obj.dataDataFInputKey)
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataF6.length > 0) {
    ucfileA1 = documentFolder.createFile(obj.myfileDataF6.setName("หนังสือสัญญา(ไฟล์หลัก)" + obj.dataDataFInputKey)).getUrl();
    const oldfile = sheetF.getRange(rowIndex + 1, 7).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetF.getRange(rowIndex + 1, 7).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetF.getRange(rowIndex + 1, 2).setValue(obj.dataDataFInput1);
  sheetF.getRange(rowIndex + 1, 3).setValue(obj.dataDataFInput2);
  sheetF.getRange(rowIndex + 1, 4).setValue(obj.dataDataFInput3);
  sheetF.getRange(rowIndex + 1, 5).setValue(obj.dataDataFInput4);
  sheetF.getRange(rowIndex + 1, 6).setValue(obj.dataDataFInput5);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
              '\n📍 เลขที่: ' + obj.dataDataFInputKey +
              '\n📍 ลงวันที่: ' + obj.dataDataFInput2 +
              '\n📝 เรื่อง: ' + obj.dataDataFInput3 +
              '\n💻 หน่วยงาน: ' + obj.dataDataFInput4 +
              '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataFInput5 +
              '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });
  
  return sheetF.getRange("A2:G" + sheetF.getLastRow()).getValues();
}

const delRecDataF = (record) =>  {
  const sheetF = SpreadsheetApp.openById(sheetDataDEF).getSheetByName('DataF');
  const data = sheetF.getDataRange().getDisplayValues()
  const iddataF = data.map(r=>r[0])
  var rowIndex = iddataF.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetF.getRange(rowIndex + 1, 7).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetF.deleteRow(rowIndex + 1);
  }
}