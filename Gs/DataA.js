// const getDataA = () => {
//   const sheetA = SpreadsheetApp.openById(sheetDataA);
//   const sheet = sheetA.getSheetByName("DataA"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataA = (userName, userStatus) => {
  const sheetA = SpreadsheetApp.openById(sheetDataA);
  const sheet = sheetA.getSheetByName("DataA");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[10] === userName);
  }
};

function generateIDDataA() {
  const sheetA = SpreadsheetApp.openById(sheetDataA);
  const sheet = sheetA.getSheetByName("DataA"); 
  const lastRowID = sheet.getLastRow();
  const codeIDDataA = generateIDDataAFromRow(lastRowID);
  return codeIDDataA;
}

function generateIDDataAFromRow(currentIDDataA) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataA.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const checkDataA = (value) => {
  var sheetName = "DataA";
  var columnsToCheck = ["F"]; 
  var sheet = SpreadsheetApp.openById(sheetDataA).getSheetByName(sheetName);

  for (var col of columnsToCheck) {
    var data = sheet.getRange(col + "1:" + col + sheet.getLastRow()).getValues().flat();
    if (data.some(cell => cell === value)) {
      return { isDuplicate: true };
    }
  }
  return { isDuplicate: false };
}

const addDataA = (obj) => {
  const sheetA = SpreadsheetApp.openById(sheetDataA).getSheetByName('DataA');
  const lastRowID = sheetA.getLastRow();
  const codeIDDataA = generateIDDataAFromRow(lastRowID);
  const documentFolder = DriveApp.getFolderById(idfolder);

  const createFileAndReturnUrl = (base64Data, namePrefix) => {
    if (!base64Data) return "";
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, 'application/pdf', namePrefix + codeIDDataA);
    return documentFolder.createFile(blob).getUrl();
  };

  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataA11, "หนังสือภายนอก(ไฟล์หลัก)") || "";
  const ucfileA2 = createFileAndReturnUrl(obj.myfileDataA12, "หนังสือภายนอก(เอกสารแนบ1)") || "";
  const ucfileA3 = createFileAndReturnUrl(obj.myfileDataA13, "หนังสือภายนอก(เอกสารแนบ2)") || "";

  const d = new Date();
  const curdate = d.getDate();
  const monthcur = d.getMonth();
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate + ' ' + monthThai + ' ' + yearcur;

  const date = obj.dataDataAInput6.split("-");
  const dateTH = Number(date[2]);
  const mounthText = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const mounthThai = mounthText[Number(date[1])];
  const yearThai = Number(date[0]) + 543;
  const dateThai = dateTH + ' ' + mounthThai + ' ' + yearThai;

  const rowData = ["'" + codeIDDataA, "รอดำเนินการ", obj.dataDataAInput2, obj.dataDataAInput3, newdateThai, obj.dataDataAInput5, dateThai, obj.dataDataAInput7, obj.dataDataAInput8, obj.dataDataAInput9, obj.dataDataAInput10, ucfileA1, ucfileA2, ucfileA3];
  sheetA.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const tinyurlA2 = ucfileA2 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA2);
  const tinyurlA3 = ucfileA3 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA3);

  const msg = "I-OFFICE แจ้งเตือน" +
    '\n📍 เลขที่: ' + codeIDDataA +
    '\n🚀 ความเร่งด่วน: ' + obj.dataDataAInput2 +
    '\n🧑 ชั้นความลับ: ' + obj.dataDataAInput3 +
    '\n📍 เลขที่หนังสือ: ' + obj.dataDataAInput5 +
    '\n📍 ลงวันที่: ' + dateThai +
    '\n🏦 จาก: ' + obj.dataDataAInput7 +
    '\n🧑 ถึง: ' + obj.dataDataAInput8 +
    '\n📝 เรื่อง: ' + obj.dataDataAInput9 +
    '\n📆 วันที่รับ: ' + newdateThai +
    '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataAInput10 +
    '\n📁 ไฟล์: ' + tinyurlA1 +
    '\n📁 เอกสารแนบ1: ' + tinyurlA2 +
    '\n📁 เอกสารแนบ2: ' + tinyurlA3;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetA.getRange("A2:N" + sheetA.getLastRow()).getValues();
}

const upDateDataA = (obj) => {
  const sheetA = SpreadsheetApp.openById(sheetDataA).getSheetByName('DataA');
  const data = sheetA.getDataRange().getDisplayValues();
  const iddataA = data.map(r => r[0]);
  const rowIndex = iddataA.indexOf(obj.dataDataAInputKey);
  const documentFolder = DriveApp.getFolderById(idfolder);

  const createFileAndReturnUrl = (base64Data, namePrefix) => {
    if (!base64Data) return "";
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, 'application/pdf', namePrefix + obj.dataDataAInputKey);
    return documentFolder.createFile(blob).getUrl();
  };

  var ucfileA1 = "";
  var ucfileA2 = "";
  var ucfileA3 = "";

  if (obj.myfileDataA11 && obj.myfileDataA11.length > 0) {
    ucfileA1 = createFileAndReturnUrl(obj.myfileDataA11, "หนังสือภายนอก(ไฟล์หลัก)");
    const oldfile = sheetA.getRange(rowIndex + 1, 12).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetA.getRange(rowIndex + 1, 12).setValue(ucfileA1);
  }

  if (obj.myfileDataA12 && obj.myfileDataA12.length > 0) {
    ucfileA2 = createFileAndReturnUrl(obj.myfileDataA12, "หนังสือภายนอก(เอกสารแนบ1)");
    const oldfile = sheetA.getRange(rowIndex + 1, 13).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetA.getRange(rowIndex + 1, 13).setValue(ucfileA2);
  }

  if (obj.myfileDataA13 && obj.myfileDataA13.length > 0) {
    ucfileA3 = createFileAndReturnUrl(obj.myfileDataA13, "หนังสือภายนอก(เอกสารแนบ2)");
    const oldfile = sheetA.getRange(rowIndex + 1, 14).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetA.getRange(rowIndex + 1, 14).setValue(ucfileA3);
  }

  if (rowIndex > -1) {
    sheetA.getRange(rowIndex + 1, 2).setValue(obj.dataDataAInput1);
    sheetA.getRange(rowIndex + 1, 3).setValue(obj.dataDataAInput2);
    sheetA.getRange(rowIndex + 1, 4).setValue(obj.dataDataAInput3);
    sheetA.getRange(rowIndex + 1, 5).setValue(obj.dataDataAInput4);
    sheetA.getRange(rowIndex + 1, 6).setValue(obj.dataDataAInput5);
    sheetA.getRange(rowIndex + 1, 7).setValue(obj.dataDataAInput6);
    sheetA.getRange(rowIndex + 1, 8).setValue(obj.dataDataAInput7);
    sheetA.getRange(rowIndex + 1, 9).setValue(obj.dataDataAInput8);
    sheetA.getRange(rowIndex + 1, 10).setValue(obj.dataDataAInput9);
    sheetA.getRange(rowIndex + 1, 11).setValue(obj.dataDataAInput10);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const tinyurlA2 = ucfileA2 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA2);
  const tinyurlA3 = ucfileA3 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA3);

  const msg = "I-OFFICE แจ้งเตือน" +
    '\n📍 เลขที่: ' + obj.dataDataAInputKey +
    '\n🚀 ความเร่งด่วน: ' + obj.dataDataAInput2 +
    '\n🧑 ชั้นความลับ: ' + obj.dataDataAInput3 +
    '\n📍 เลขที่หนังสือ: ' + obj.dataDataAInput5 +
    '\n📍 ลงวันที่: ' + obj.dataDataAInput6 +
    '\n🏦 จาก: ' + obj.dataDataAInput7 +
    '\n🧑 ถึง: ' + obj.dataDataAInput8 +
    '\n📝 เรื่อง: ' + obj.dataDataAInput9 +
    '\n📆 วันที่รับ: ' + obj.dataDataAInput4 +
    '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataAInput10 +
    '\n📁 ไฟล์: ' + tinyurlA1 +
    '\n📁 เอกสารแนบ1: ' + tinyurlA2 +
    '\n📁 เอกสารแนบ2: ' + tinyurlA3;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetA.getRange("A2:N" + sheetA.getLastRow()).getValues();
}

const delRecDataA = (record) =>  {
  const sheetA = SpreadsheetApp.openById(sheetDataA).getSheetByName('DataA');
  const data = sheetA.getDataRange().getDisplayValues()
  const iddataA = data.map(r=>r[0])
  var rowIndex = iddataA.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetA.getRange(rowIndex + 1, 12).getValue();
    const ucfileA2 = sheetA.getRange(rowIndex + 1, 13).getValue();
    const ucfileA3 = sheetA.getRange(rowIndex + 1, 14).getValue(); 
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    if(ucfileA2 !=""){
      DriveApp.getFileById(ucfileA2.split('/')[5]).setTrashed(true)
    }
    if(ucfileA3 !=""){
      DriveApp.getFileById(ucfileA3.split('/')[5]).setTrashed(true)
    }
    sheetA.deleteRow(rowIndex + 1);
  }
}