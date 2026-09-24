// const getDataC = () => {
//   const sheetC = SpreadsheetApp.openById(sheetDataC);
//   const sheet = sheetC.getSheetByName("DataC"); 
//   const data = sheet.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataC = (userName, userStatus) => {
  const sheetC = SpreadsheetApp.openById(sheetDataC);
  const sheet = sheetC.getSheetByName("DataC");
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[7] === userName);
  }
};

function generateIDDataC(currentIDDataC) {
  const today = new Date();
  const thaiYear = today.getFullYear() + 543;
  const number = currentIDDataC.toString().padStart(3, '0');
  return `${number}/${thaiYear}`;
}

const addDataC = (obj) => {
  const sheetC = SpreadsheetApp.openById(sheetDataC).getSheetByName('DataC');
  const lastRowID = sheetC.getLastRow();
  var codeIDDataC = generateIDDataC(lastRowID); 
  const documentFolder = DriveApp.getFolderById(idfolder);
  const createFileAndReturnUrl = (fileData, namePrefix) => {
    return fileData.length > 0 ? documentFolder.createFile(fileData.setName(namePrefix + codeIDDataC)).getUrl() : "";
  };
  const ucfileA1 = createFileAndReturnUrl(obj.myfileDataC8, "หนังสือส่งออก(ไฟล์หลัก)")|| "";

  const d = new Date(); 
  const curdate=d.getDate(); 
  const monthcur=d.getMonth(); 
  const yearcur =d.getFullYear()+543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
  const monthThai = monthCut[monthcur];
  const newdateThai = curdate+' '+monthThai+' '+yearcur;

  const rowData = ["'"+codeIDDataC, "รอดำเนินการ", newdateThai, obj.dataDataCInput3, obj.dataDataCInput4, obj.dataDataCInput5, obj.dataDataCInput6, obj.dataDataCInput7, ucfileA1];
  sheetC.appendRow(rowData);

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
            '\n📍 เลขที่: ' + codeIDDataC+
            '\n📍 ลงวันที่: ' + newdateThai +
            '\n🏦 จาก: ' + obj.dataDataCInput3 +
            '\n🧑 ถึง: ' + obj.dataDataCInput4 +
            '\n📝 เรื่อง: ' + obj.dataDataCInput5 +
            '\n💻 หน่วยงาน: ' + obj.dataDataCInput6 +
            '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataCInput7 +
            '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetC.getRange("A2:I" + sheetC.getLastRow()).getValues();
}

const upDateDataC = (obj) => {
  const sheetC = SpreadsheetApp.openById(sheetDataC).getSheetByName('DataC');
  const data = sheetC.getDataRange().getDisplayValues()
  const iddataC = data.map(r=>r[0])
  const rowIndex = iddataC.indexOf(obj.dataDataCInputKey)
  const documentFolder = DriveApp.getFolderById(idfolder);

  var ucfileA1 = "";
  if (obj.myfileDataC8.length > 0) {
    ucfileA1 = documentFolder.createFile(obj.myfileDataC8.setName("หนังสือส่งออก(ไฟล์หลัก)" + obj.dataDataCInputKey)).getUrl();
    const oldfile = sheetC.getRange(rowIndex + 1, 9).getValue().split('/')[5];
    if (oldfile) {
      DriveApp.getFileById(oldfile).setTrashed(true);
    }
    sheetC.getRange(rowIndex + 1, 9).setValue(ucfileA1);
  }

  if(rowIndex > -1){
  sheetC.getRange(rowIndex + 1, 2).setValue(obj.dataDataCInput1);
  sheetC.getRange(rowIndex + 1, 3).setValue(obj.dataDataCInput2);
  sheetC.getRange(rowIndex + 1, 4).setValue(obj.dataDataCInput3);
  sheetC.getRange(rowIndex + 1, 5).setValue(obj.dataDataCInput4);
  sheetC.getRange(rowIndex + 1, 6).setValue(obj.dataDataCInput5);
  sheetC.getRange(rowIndex + 1, 7).setValue(obj.dataDataCInput6);
  sheetC.getRange(rowIndex + 1, 8).setValue(obj.dataDataCInput7);
  }

  const tinyurlA1 = ucfileA1 === "" ? "ไม่ได้แนบไฟล์" : shortenURL(ucfileA1);

  const msg = "I-OFFICE แจ้งเตือน" +
            '\n📍 เลขที่: ' + obj.dataDataCInputKey+
            '\n📍 ลงวันที่: ' + obj.dataDataCInput2 +
            '\n🏦 จาก: ' + obj.dataDataCInput3 +
            '\n🧑 ถึง: ' + obj.dataDataCInput4 +
            '\n📝 เรื่อง: ' + obj.dataDataCInput5 +
            '\n💻 หน่วยงาน: ' + obj.dataDataCInput6 +
            '\n💻 ผู้ลงทะเบียน: ' + obj.dataDataCInput7 +
            '\n📁 ไฟล์: ' + tinyurlA1;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

  return sheetC.getRange("A2:I" + sheetC.getLastRow()).getValues();
}

const delRecDataC = (record) =>  {
  const sheetC = SpreadsheetApp.openById(sheetDataC).getSheetByName('DataC');
  const data = sheetC.getDataRange().getDisplayValues()
  const iddataC = data.map(r=>r[0])
  var rowIndex = iddataC.indexOf(record);
  if (rowIndex > -1) {
    const ucfileA1 = sheetC.getRange(rowIndex + 1, 9).getValue();
    if(ucfileA1 !=""){
      DriveApp.getFileById(ucfileA1.split('/')[5]).setTrashed(true)
    }
    sheetC.deleteRow(rowIndex + 1);
  }
}