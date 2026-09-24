const getBookType = (dataType) => {
  switch (dataType) {
    case 'DataA':
      return "หนังสือรับ";
    case 'DataB':
      return "หนังสือรับ";
    case 'DataD':
      return "บันทึก";
    case 'DataE':
      return "คำสั่ง";
    default:
      return "ไม่ทราบประเภท";
  }
}

// const getDataAlSA = () => {
//   const sheet = SpreadsheetApp.openById(sheetDataSA);
//   const sheetset = sheet.getSheetByName("DataSA"); 
//   const data = sheetset.getDataRange().getDisplayValues().slice(1);
//   //Logger.log(data)
//   return data;
// }

const getDataAlSA = () => {
  const sheet = SpreadsheetApp.openById(sheetDataSA);
  const sheetset = sheet.getSheetByName("DataSA"); 
  const dataA = sheetset.getDataRange().getDisplayValues().slice(1).map(row => [...row, 'ส่งภายใน']);

  const combinedData = [...dataA];
  const filteredData = combinedData.filter(row => row[2] === "รอตรวจสอบ" || row[2] === "ติดตาม");

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  const userData = userSheet.getDataRange().getValues();
  const userMap = createUserMap(userData);

  const uniqueData = filterUniqueById(filteredData);
  const latestData = uniqueData.slice(0, 5);

  Logger.log(latestData);

  return latestData.map(row => ({
    status: row[1],
    idstatus: row[2],
    name: row[6],
    type: row[row.length - 1],
    img: userMap[row[6]] || 'https://cdn.jsdelivr.net/gh/EPICCODING17/image/Logo-EicCoding.png',
    date: row[3],
    department: row[7]
  }));
}

const createUserMap = (userData) => {
  const userMap = {};
  userData.forEach(row => {
    const name = row[3];
    const img = row[6];
    userMap[name] = img;
  });
  return userMap;
}

const filterUniqueById = (data) => {
  const seen = new Set();
  return data.filter(row => {
    const id = row[3];
    if (seen.has(id)) {
      return false;
    }
    seen.add(id);
    return true;
  });
}

const getDataSA = (userName, userStatus) => {
  const sheet = SpreadsheetApp.openById(sheetDataSA);
  const sheetset = sheet.getSheetByName("DataSA"); 
  const data = sheetset.getDataRange().getDisplayValues().slice(1);
  
  if (userStatus === "SuperAdmin" || userStatus === "Admin") {
    return data;
  } else {
    return data.filter(row => row[6] === userName);
  }
};

const generateIDDataSA = (currentIDDataSA) => {
  var prefix = 'SA-';
  var today = new Date();
  var thaiYear = (today.getFullYear() + 543).toString().slice(-2);
  var number = currentIDDataSA.toString().padStart(4, '0');
  return prefix + thaiYear + number;
}

const sendAddDataSA = (obj) => {
  const sheetSA = SpreadsheetApp.openById(sheetDataSA).getSheetByName('DataSA');
  const lastRowID = sheetSA.getLastRow();
  const currentIDDataSA = (lastRowID > 1) ? (lastRowID - 1) : 0;
  const d = new Date();
  const curdate = d.getDate();
  const monthcur = d.getMonth();
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateSAThai = curdate + ' ' + monthThai + ' ' + yearcur;
  const selectedUsers = JSON.parse(obj.selectedUsers);
  const bookType = getBookType(obj.dataType);
  selectedUsers.forEach((userTK, index) => {
  const codeIDDataSA = generateIDDataSA(currentIDDataSA + index + 1);
  const rowData = [codeIDDataSA, bookType, "รอตรวจสอบ", newdateSAThai, obj.dataInputKey, obj.dataInput1, userTK.fullname, userTK.department, obj.dataInput2, obj.fileData1, obj.fileData2, obj.fileData3,"", "", "", "0"];
  sheetSA.appendRow(rowData);

  const shortfileData1 = obj.fileData1 ? shortenURL(obj.fileData1) : 'ไม่ได้แนบไฟล์';
  const shortfileData2 = obj.fileData2 ? shortenURL(obj.fileData2) : 'ไม่ได้แนบไฟล์';
  const shortfileData3 = obj.fileData3 ? shortenURL(obj.fileData3) : 'ไม่ได้แนบไฟล์';
  const msg = "I-OFFICE แจ้งเตือนหนังสือส่ง" +
                  '\n📍 เลขที่: ' + codeIDDataSA +
                  '\n📍 วันที่ส่ง: ' + newdateSAThai +
                  '\n📍 ประเภท: ' + bookType +
                  '\n📍 สถานะ: ' + "รอตรวจสอบ" +
                  '\n📍 เลขที่ ' + obj.dataInputKey +
                  '\n📍 เรื่อง: ' + obj.dataInput1 +
                  '\n📍 ชื่อผู้รับ: ' + userTK.fullname +
                  '\n📍 หน่วยงาน: ' + userTK.department +
                  '\n📍 รายละเอียด: ' + obj.dataInput2 +
                  '\n📁 ไฟล์: ' + shortfileData1 +
                  '\n📁 เอกสารแนบ1: ' + shortfileData2 +
                  '\n📁 เอกสารแนบ2: ' + shortfileData3;

    const activeNotifications = getNotificationSettings();
    
    activeNotifications.forEach(setting => {
      sendNotify(msg, setting);
    });

  });
  return sheetSA.getRange("A2:L" + sheetSA.getLastRow()).getValues();
}

const followRecord = (record, receiverName) => {
  const sheet = SpreadsheetApp.openById(sheetDataSA).getSheetByName("DataSA");
  const data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === record) {
      data[i][2] = "ติดตาม";
      data[i][15] = parseInt(data[i][15] || 0) + 1;
      sheet.getRange(i + 1, 3).setValue(data[i][2]);
      sheet.getRange(i + 1, 16).setValue(data[i][15]);
      break;
    }
  }

  const msg = "I-OFFICE แจ้งเตือนติดตามหนังสือส่งภายใน" +
              '\n📍 เลขที่ส่ง: ' + record +
              '\n📍 ชื่อผู้รับ: ' + receiverName;

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });
}

const sigSaveDataSA = (saveData) => {
  const { record, receiverName, sigSA, remark, action } = saveData;
  const sheet = SpreadsheetApp.openById(sheetDataSA).getSheetByName("DataSA");
  const data = sheet.getDataRange().getValues();
  const d = new Date();
  const curdate = d.getDate();
  const monthcur = d.getMonth();
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateSAThai = curdate + ' ' + monthThai + ' ' + yearcur;

  let type, number, subject, department, details;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == record) {
      const status = action === 'confirm' ? "เสร็จสิ้น" : "ตีกลับ";
      sheet.getRange(i + 1, 3).setValue(status);
      sheet.getRange(i + 1, 13).setValue(sigSA);
      sheet.getRange(i + 1, 14).setValue(newdateSAThai);
      sheet.getRange(i + 1, 15).setValue(remark ? remark : '');
      type = data[i][1];
      number = data[i][4];
      subject = data[i][5];
      department = data[i][7];
      details = data[i][8];
      break;
    }
  }
    const msg = "I-OFFICE แจ้งเตือนรับหนังสือส่งภายใน" +
                '\n📍 เลขที่ส่ง: ' + record +
                '\n📍 ประเภท: ' + type +
                '\n📍 สถานะ: ' + (action === 'confirm' ? 'เสร็จสิ้น' : 'ตีกลับ') +
                '\n📍 เลขที่เอกสาร: ' + number +
                '\n📍 เรื่อง: ' + subject +
                '\n📍 ชื่อผู้รับ: ' + receiverName +
                '\n📍 หน่วยงาน: ' + department +
                '\n📍 รายละเอียด: ' + details +
                (remark ? '\n📍 หมายเหตุ: ' + remark : '');

  const activeNotifications = getNotificationSettings();
  
  activeNotifications.forEach(setting => {
    sendNotify(msg, setting);
  });

}

const getDataSB = () => {
  const sheet = SpreadsheetApp.openById(sheetDataSB);
  const sheetset = sheet.getSheetByName("DataSB"); 
  const data = sheetset.getDataRange().getDisplayValues().slice(1);
  //Logger.log(data)
  return data;
}

const generateIDDataSB = (currentIDDataSB) => {
  var prefix = 'SB-';
  var today = new Date();
  var thaiYear = (today.getFullYear() + 543).toString().slice(-2);
  var number = currentIDDataSB.toString().padStart(4, '0');
  return prefix + thaiYear + number;
}

const sendDataAgency = (obj) => {
  const sheetSB = SpreadsheetApp.openById(sheetDataSB).getSheetByName('DataSB');
  const lastRowID = sheetSB.getLastRow();
  const currentIDDataSB = (lastRowID > 1) ? (lastRowID - 1) : 0;
  const d = new Date();
  const curdate = d.getDate();
  const monthcur = d.getMonth();
  const yearcur = d.getFullYear() + 543;
  const monthCut = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const monthThai = monthCut[monthcur];
  const newdateSBThai = curdate + ' ' + monthThai + ' ' + yearcur;

  const selectedAgency = JSON.parse(obj.selectedAgency);
  selectedAgency.forEach((rowAge, index) => {
    const codeIDDataSB = generateIDDataSB(currentIDDataSB + index + 1);
    const rowData = [codeIDDataSB, newdateSBThai, obj.ageInputKey, obj.ageInput1, obj.ageInput2, obj.ageInput3, rowAge.fullname, rowAge.email, obj.ageInput4, obj.ageInput5, obj.ageInput6, obj.fileDataAge1];
    sheetSB.appendRow(rowData);
    const shortfileAge1 = obj.fileDataAge1 ? shortenURL(obj.fileDataAge1) : 'ไม่ได้แนบไฟล์';
    const emailSubject = "I-OFFICE ระบสารบรรณ ออนไลน์";
    const emailBody = `
     <div style="border: 1px solid #ddd; border-radius: 10px; padding: 20px; font-family: Arial, sans-serif; max-width: 600px; margin: auto;">
        <img src="${logoUrl}" style="width: 50px; border-radius: 10px 10px 0 0;" alt="...">
        <div style="padding: 20px;">
          <h5 style="text-align: center; color: #4CAF50;">I-OFFICE ระบสารบรรณ ออนไลน์</h5>
          <p><strong>📍 เลขที่ทำรายการ:</strong> ${codeIDDataSB}</p>
          <p><strong>📍 วันที่จัดส่ง:</strong> ${newdateSBThai}</p>
          <p><strong>📍 จาก:</strong> ${obj.ageInput1}</p>
          <p><strong>📍 ถึง:</strong> ${obj.ageInput2}</p>
          <p><strong>📍 เลขที่:</strong> ${obj.ageInputKey}</p>
          <p><strong>📍 เรื่อง:</strong> ${obj.ageInput3}</p>
          <p><strong>📍 หน่วยงานที่รับ:</strong> ${rowAge.fullname}</p>
          <p><strong>📍 อีเมล:</strong> ${rowAge.email}</p>
          <p><strong>📍 รายละเอียด:</strong> ${obj.ageInput4}</p>
          <p><strong>📍 การแจ้ง:</strong> ${obj.ageInput5}</p>
          <p><strong>📍 ผู้ส่ง:</strong> ${obj.ageInput6}</p>
          <a href="${shortfileAge1}" style="display: inline-block; padding: 10px 20px; margin-top: 20px; background-color: #007BFF; color: white; text-decoration: none; border-radius: 5px;">เปิดเอกสารแนบ</a>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: rowAge.email,
      subject: emailSubject,
      htmlBody: emailBody
    });
  });
  return sheetSB.getRange("A2:J" + sheetSB.getLastRow()).getValues();
};