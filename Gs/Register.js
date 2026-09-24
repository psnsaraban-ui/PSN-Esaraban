var sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users')
var data = sheetUsers.getDataRange().getDisplayValues()
var iddataLU = data.map(r=>r[0])

const getDataLU = () => {
  var data = sheetUsers.getDataRange().getDisplayValues().slice(1)
  //Logger.log(data)
  return data
}


const getDataLogUS = (statususer) =>{
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LogUser'); 
  var data = ss.getDataRange().getDisplayValues()
  var resultlogus
  if(statususer == 'admin'){
    resultlogus = data.slice(1)
  }else{
    resultlogus = data.filter(r=>r[0] == statususer)
  }

  //Logger.log(resultlogus)
  return resultlogus
}

const findRecordU = (recordu) => {
  var data = sheetUsers.getDataRange().getValues()
  var iddataLU = data.map(r => r[0])
  var index = iddataLU.indexOf(recordu) 
  if(index > -1){
    data = sheetUsers.getRange(index+1,1,1,sheetUsers.getLastColumn()).getDisplayValues()[0]
  }
  //Logger.log(data)
  return data
}

const checkFnameUser = (value) => {
  var sheetName = "Users";
  var columnToCheck = "B"; 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data = sheet.getRange(columnToCheck + "1:" + columnToCheck + sheet.getLastRow()).getValues();
  for (var i = 0; i < data.length; i++) {
        if (data[i][0] === value) {
        return { isDuplicate: true };
    }
  }
  return { isDuplicate: false };
}

const generateIDMember = (currentIDNumber) => {
  var prefix = 'USER-';
  var paddingSize = 3;
  var number = currentIDNumber.toString();
  while (number.length < paddingSize) {
    number = '0' + number;
  }
  return prefix + number;
}

const saveUser = (obj) => {
  const sheetUsers = SpreadsheetApp.getActive().getSheetByName('Users');
  const lastRowID = sheetUsers.getLastRow();
  var codeUserIDMember = generateIDMember(lastRowID);
  var folder = DriveApp.getFolderById(imageFolder);
  var profileUrl = "";

  if (obj.check !== "") {
    var datafile = Utilities.base64Decode(obj.imageDataUrlA.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    profileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
  } else {
    profileUrl = obj.profile;
  }
      
      sheetUsers.appendRow(["'" + codeUserIDMember,
                    "'"+obj.registerData4, 
                    "'"+obj.registerData5,
                    "'"+obj.registerData3,  
                    obj.registerData1,
                    obj.registerData2, 
                    profileUrl,         
                    ]);

  var data = sheetUsers.getDataRange().getDisplayValues();
  return data;
}

const editUser = (obj) => {
  var rowIndex = iddataLU.indexOf(obj.registerDataID)
  var folder = DriveApp.getFolderById(imageFolder);
  var profileUrl = "";

  if (obj.check !== "") {
    if (obj.imageDataUrlA.length > 0) {  
    var datafile = Utilities.base64Decode(obj.imageDataUrlA.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    profileUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    let oldprofile = sheetUsers.getRange(rowIndex+1,7).getValue().split('/d/')[1]
    if(oldprofile!=undefined){
      DriveApp.getFileById(oldprofile).setTrashed(true)
    }
  }
  }else {
    profileUrl = obj.profile;
  }
  if (rowIndex > -1) {
    sheetUsers.getRange(rowIndex+1,1,1,8).setValues([[obj.registerDataID, "'"+obj.registerData4, "'"+obj.registerData5, "'"+obj.registerData3, obj.registerData1, obj.registerData2, profileUrl]]);
  }
   var data = sheetUsers.getRange(rowIndex+1,1,1,sheetUsers.getLastColumn()).getDisplayValues()[0]
   return data;
}

const delRecordU = (record) =>  {
  var rowIndex = iddataLU.indexOf(record);
  if (rowIndex > -1) {
    const fileDlUser = sheetUsers.getRange(rowIndex + 1, 7).getValue();
    const sigDlUser = sheetUsers.getRange(rowIndex + 1, 8).getValue(); 

    if (fileDlUser.includes("https://lh3.googleusercontent.com/d/")) {
      const fileId = fileDlUser.split('/d/')[1];
      DriveApp.getFileById(fileId).setTrashed(true);
    }

    if (sigDlUser.includes("https://lh3.googleusercontent.com/d/")) {
      const fileId = sigDlUser.split('/d/')[1];
      DriveApp.getFileById(fileId).setTrashed(true);
    }

    sheetUsers.deleteRow(rowIndex + 1);
  }
}

function updateProfile(obj) {
  const sheetUsers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheetUsers.getDataRange().getValues();
  var folder = DriveApp.getFolderById(imageFolder);
  var profileUrl = "";

  if (obj.check !== "") {
    var datafile = Utilities.base64Decode(obj.imageDataUrl.split(',')[1]);
    var blob = Utilities.newBlob(datafile, obj.filetype, obj.filename);
    var file = folder.createFile(blob);
    var fileId = file.getId();
    profileUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    for (let i = 1; i < data.length; i++) {
      const userValue = data[i][1];
      if (userValue === obj.codeName) {
        const oldprofileValue = data[i][6];
        if (oldprofileValue && oldprofileValue.startsWith("https://lh3.googleusercontent.com/d/") && oldprofileValue !== "https://lh3.googleusercontent.com/d/") {
          let oldprofile = oldprofileValue.split('/d/')[1];
          if (oldprofile) {
            DriveApp.getFileById(oldprofile).setTrashed(true);
          }
        }
        sheetUsers.getRange(i + 1, 7).setValue(profileUrl);
        break;
      }
    }
    return profileUrl;
  } else {
    profileUrl = obj.profileNew;
    return profileUrl;
  }
}

function saveUploadSig(obj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  var folderSig = DriveApp.getFolderById(sigFolder);

  if (obj.signatureCanvas === "") {
    for (let i = 1; i < data.length; i++) {
      const userValue = data[i][1];
      if (userValue === obj.codeName) {
        const oldUrlSig = data[i][7];
        if (oldUrlSig) {
          const oldSig = oldUrlSig.split('/d/')[1];
          const oldFile = DriveApp.getFileById(oldSig);
          oldFile.setTrashed(true);
          sheet.getRange(i + 1, 8).setValue("");
          break;
        }
      }
    }
    return "";
  } else {
    var datafile2 = Utilities.base64Decode(obj.signatureCanvas.split(',')[1]);
    var blob2 = Utilities.newBlob(datafile2, obj.filetype, obj.filename);
    var fileSig = folderSig.createFile(blob2);
    var sig = fileSig.getId();
    var urlsig = "https://lh3.googleusercontent.com/d/" + sig;

    for (let i = 1; i < data.length; i++) {
      const userValue = data[i][1];
      if (userValue === obj.codeName) {
        const oldUrlSig = data[i][7];
        if (oldUrlSig) {
          const oldSig = oldUrlSig.split('/d/')[1];
          const oldFile = DriveApp.getFileById(oldSig);
          oldFile.setTrashed(true);
        }
        sheet.getRange(i + 1, 8).setValue(urlsig);
        break;
      }
    }
    return urlsig;
  }
}

function ChangePasswordMember(obj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users"); 
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const userValue = data[i][0];
    if (userValue === obj.code) {
      sheet.getRange(i + 1, 3).setValue("'" + obj.password);
      break;
    }
  }
}