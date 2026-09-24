// const getDataB = () => {
//   const sheetB = SpreadsheetApp.openById(sheetDataB);
//   const sheet = sheetB.getSheetByName("DataB"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataB = (userName, userStatus) => {
  const sheetB = SpreadsheetApp.openById(sheetDataB);
  const sheet = sheetB.getSheetByName("DataB");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[8] === userName);
  }
};

function generateIDDataB(currentIDDataB) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataB.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const addDataB = (obj) => {
  const sheetB = SpreadsheetApp.openById(sheetDataB).getSheetByName('DataB');
  const lastRowID = sheetB.getLastRow();
  var codeIDDataB = generateIDDataB(lastRowID); 
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => createPdfFileFromBase64(documentFolder, fileData, namePrefix + codeIDDataB);
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataB9, "หนังสือภายใน(ไฟล์หลัก)")|| "";
  const ucfileA2 = createFileAndReturnUrl(obj.myfileDataB10, "หนังสือภายใน(เอกสารแนบ1)")|| "";

  const d = new Date(); 
  const curdate=d.getDate(); 
  const monthcur=d.getMonth(); 
  const yearcur =d.getFullYear()+543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate+' '+monthThai+' '+yearcur;

  const date = obj.dataDataBInput6.split("-") 
  const dateTH = Number(date[2]);
  const mounthText = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const mounthThai = mounthText[Number(date[1])];
  const yearThai = Number(date[0])+543
  const dateThai = dateTH+' '+mounthThai+' '+yearThai

  const rowData = ["'"+codeIDDataB, "รอดำเนินการ", obj.dataDataBInput2, obj.dataDataBInput3, newdateThai, obj.dataDataBInput5, dateThai, obj.dataDataBInput7, obj.dataDataBInput8, ucfileA1, ucfileA2];
  sheetB.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const tinyurlA2 = ucfileA2 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA2);

  const msg = "I-OFFICE แจ้งเตือน" +
            '\n📍 เลขที่: ' + codeIDDataB+
            '\n🚀 ความเร่งด่วน: ' + obj.dataDataBInput2 +
            '\n🧑 ชั้นความลับ: ' + obj.dataDataBInput3 +
            '\n📍 เลขที่หนังสือ: ' + obj.dataDataBInput5 +
            '\n📍 ลงวันที่: ' + dateThai +
            '\n📝 เรื่อง: ' + obj.dataDataBInput7 +
            '\n📆 วันที่รับ: ' + newdateThai +
            '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataBInput8 +
            '\n📁 ไฟล์: ' + tinyurlA1 +
            '\n📁 เอกสารแนบ1: ' + tinyurlA2

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetB.getRange("A2:K" + sheetB.getLastRow()).getValues();
}

const upDateDataB = (obj) => {
  const sheetB = SpreadsheetApp.openById(sheetDataB).getSheetByName('DataB');
  const data = sheetB.getDataRange().getDisplayValues()
  const iddataB = data.map(r=>r[0])
  const rowIndex = iddataB.indexOf(obj.dataDataBInputKey)
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  var ucfileA2 = "";

  if (obj.myfileDataB9) {
    ucfileA1 = createPdfFileFromBase64(documentFolder, obj.myfileDataB9, "หนังสือภายใน(ไฟล์หลัก)" + obj.dataDataBInputKey);
    const oldfile = sheetB.getRange(rowIndex + 1, 10).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetB.getRange(rowIndex + 1, 10).setValue(ucfileA1);
  }

  if (obj.myfileDataB10) {
    ucfileA2 = createPdfFileFromBase64(documentFolder, obj.myfileDataB10, "หนังสือภายใน(เอกสารแนบ1)" + obj.dataDataBInputKey);
    const oldfile = sheetB.getRange(rowIndex + 1, 11).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetB.getRange(rowIndex + 1, 11).setValue(ucfileA2);
  }

  if(rowIndex > -1){
  sheetB.getRange(rowIndex + 1, 2).setValue(obj.dataDataBInput1);
  sheetB.getRange(rowIndex + 1, 3).setValue(obj.dataDataBInput2);
  sheetB.getRange(rowIndex + 1, 4).setValue(obj.dataDataBInput3);
  sheetB.getRange(rowIndex + 1, 5).setValue(obj.dataDataBInput4);
  sheetB.getRange(rowIndex + 1, 6).setValue(obj.dataDataBInput5);
  sheetB.getRange(rowIndex + 1, 7).setValue(obj.dataDataBInput6);
  sheetB.getRange(rowIndex + 1, 8).setValue(obj.dataDataBInput7);
  sheetB.getRange(rowIndex + 1, 9).setValue(obj.dataDataBInput8);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);
  const tinyurlA2 = ucfileA2 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA2);

  const msg = "I-OFFICE แจ้งเตือน" +
            '\n📍 เลขที่: ' + obj.dataDataBInputKey+
            '\n🚀 ความเร่งด่วน: ' + obj.dataDataBInput2 +
            '\n🧑 ชั้นความลับ: ' + obj.dataDataBInput3 +
            '\n📍 เลขที่หนังสือ: ' + obj.dataDataBInput5 +
            '\n📍 ลงวันที่: ' + obj.dataDataBInput4 +
            '\n📝 เรื่อง: ' + obj.dataDataBInput7 +
            '\n📆 วันที่รับ: ' + obj.dataDataBInput6 +
            '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataBInput8 +
            '\n📁 ไฟล์: ' + tinyurlA1 +
            '\n📁 เอกสารแนบ1: ' + tinyurlA2

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetB.getRange("A2:K" + sheetB.getLastRow()).getValues();
}

const delRecDataB = (record) =>  {
  const sheetB = SpreadsheetApp.openById(sheetDataB).getSheetByName('DataB');
  const data = sheetB.getDataRange().getDisplayValues()
  const iddataB = data.map(r=>r[0])
  var rowIndex = iddataB.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetB.getRange(rowIndex + 1, 10).getValue();
    const ucfileA2 = sheetB.getRange(rowIndex + 1, 11).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    if(ucfileA2 !=""){
      DriveApp.getFileById(ucfileA2.split('/')[5]).setTrashed(true)
    }
    sheetB.deleteRow(rowIndex + 1);
  }
}