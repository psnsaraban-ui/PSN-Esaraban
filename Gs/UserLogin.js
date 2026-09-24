function checkLogin(username, password, userIpAddress, userAgent) {
  const ws = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = ws.getDataRange().getDisplayValues();

  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android) [\d.]+/);

  const userlog = data.find(r => r[1] == username && r[2] == password);

  if (userlog) {
    const logmember = SpreadsheetApp.getActive().getSheetByName('LogUser');
    const userInfo = userlog[1];
    const logInfo = ["'" + userInfo, "เข้าสู่ระบบ", userIpAddress, (browserInfo ? browserInfo[0] : "Unknown Browser") + " " + (osInfo ? osInfo[0] : "Unknown OS"), new Date()];
    logmember.appendRow(logInfo);
  } else {
    console.log("User not found");
  }

  return userlog;
}

function appendLogRow(User, action, userIpAddress, userAgent) {
  const logmember = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogUser");
  const browserInfo = userAgent.match(/(Chrome|Safari|Firefox|Edge|Opera)\/[\d.]+/);
  const osInfo = userAgent.match(/(Windows NT|Windows|Linux|Mac OS|iOS|Android|iPhone) [\d.]+/);

  logmember.appendRow(["'" + User, action, userIpAddress, (browserInfo ? browserInfo[0] : "Unknown Browser") + " " + (osInfo ? osInfo[0] : "Unknown OS"), new Date()]);
}

function checkLogoutUser(User, userIpAddress, userAgent) {
  appendLogRow(User, "ออกจากระบบ", userIpAddress, userAgent);
}