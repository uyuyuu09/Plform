function findRow(sheet, val, col){
    var dat = sheet.getDataRange().getValues();
     
    for (var i = 1; i < dat.length; i++) {
        if (dat[i][col - 1] == val) {
            return i + 1;
        }
    }
    return 0;
}

function doGet(e) {
    var page = e.pathInfo ? e.pathInfo : "index";

    var temp = (() => {
          try {
              return HtmlService.createTemplateFromFile(page);
          } catch(e) {
              return HtmlService.createTemplateFromFile("error");
          }
    })();

    var parameter = (() => {
        try {
            return e.parameter.page;
        } catch(e) {
            return "dummy";
        }
    });

    const loginUserGmail = Session.getActiveUser().getEmail();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const member = ss.getSheetByName("名簿");
    const KBC_DB = ss.getSheetByName("部内名簿");
    try {
        var user_name = member.getRange(findRow(member,loginUserGmail,3),2).getValue();
        var class_num = member.getRange(findRow(member,loginUserGmail,3),1).getValue();
        var department = KBC_DB.getRange(findRow(KBC_DB,loginUserGmail,5),3).getValue();
    } catch {
        var user_name = "ゲストさん";
        var class_num = "9999"
        var department = "外部";
    }
    
    const adminSheet = ss.getSheetByName("admin");
    const eventName = adminSheet.getRange("C7").getValue();
    const deadLine = adminSheet.getRange("C8").getValue();
    
    temp.page = parameter;
    temp.url = ScriptApp.getService().getUrl();
    temp.gmail = loginUserGmail;
    temp.user = user_name;
    temp.class_num = class_num;
    temp.department = department;
    temp.eventName = eventName;
    temp.deadLine = deadLine;

    let res = temp
              .evaluate()
              .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
              .setTitle('放送部支援ツール「Plform」')
              .setFaviconUrl('https://drive.google.com/uc?id=1wIpqx3qVvW9ohYZCx2fFHOZCGtI1vjqC&.png')
              .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1.0');
    return res;
}

function upDateClassSheetName(row, col, newValue, classSheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(classSheetName);
    sheet.getRange(row + 7, col + 1).setValue(newValue);
}

function searchClassName(user, eventName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const db_sheet = ss.getSheetByName("団体情報記録");
     
    if(eventName.includes("体育祭")) {
        try {
            let searchClass = db_sheet.getRange(findRow(db_sheet,user,5),2).getValue();
            let sheet = ss.getSheetByName(searchClass + ".指示情報(体育祭)");
            return (sheet !== null);
        } catch {
            return false;
        }
    } else {
        try {
            let searchClass = db_sheet.getRange(findRow(db_sheet,user,5),2).getValue();
            let sheet = ss.getSheetByName(searchClass + ".指示情報");
            return (sheet !== null);
        } catch {
            return false;
        }
    }
}

function getData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const loginUserGmail = Session.getActiveUser().getEmail();
    const mem_DB = ss.getSheetByName("名簿");
    const db_sheet = ss.getSheetByName("団体情報記録");
    const adminSheet = ss.getSheetByName("admin");
    switch (arguments[0]) {
        case 'user_name':
                try {
                    var user_name = mem_DB.getRange(findRow(mem_DB, loginUserGmail, 3), 2).getValue();
                } catch {
                    var user_name = "ゲスト";
                }
                return user_name;

        case 'classData_classSheetName':
                let class_name = db_sheet.getRange(findRow(db_sheet,arguments[1],5),2).getValue();

                if(arguments[2].includes("体育祭")) {
                    let classSheetName = class_name + ".指示情報(体育祭)";
                    let mainMusicData = ss.getSheetByName(classSheetName).getRange("A17").getValue();
                    let mainMusicTimeData = ss.getSheetByName(classSheetName).getRange("E17").getValue();
                    let classMusicNameData = ss.getSheetByName(classSheetName).getRange("B8:B14").getValues();
                    let classMusicTimeData = ss.getSheetByName(classSheetName).getRange("F8:F14").getValues();
                    let musicFadeInData = ss.getSheetByName(classSheetName).getRange("H8:H14").getValues();
                    let musicFadeOutData = ss.getSheetByName(classSheetName).getRange("I8:I14").getValues();
                    return {mainMusicData, mainMusicTimeData, classMusicNameData, classMusicTimeData, classSheetName, musicFadeInData, musicFadeOutData};
                } else {
                    let classSheetName = class_name + ".指示情報";
                    let classData = ss.getSheetByName(classSheetName).getRange("A8:H27").getValues();
                    return {classData, classSheetName};
                }
        case 'AdminInfomation':
                let AdminLoginId = adminSheet.getRange("C9").getValue();
                
                if(arguments[1] === AdminLoginId) {
                    let eventName = adminSheet.getRange("C7").getValue();
                    let deadLine = adminSheet.getRange("C8").getValue();
                    let dbInfo = db_sheet.getDataRange().getValues().slice(1);
                    let msg = "HTTPステータス : 200 OK<br />"
                    return {msg, eventName, deadLine, dbInfo}
                } else {
                    let eventName = "";
                    let deadLine = "";
                    let dbInfo = "";
                    let msg = "HTTPステータス : 401 Unauthorized<br />"
                    return {msg, eventName, deadLine, dbInfo}
                }
    }
}

function sendData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const db_sheet = ss.getSheetByName("団体情報記録");
    const adminSheet = ss.getSheetByName("admin");
    switch (arguments[0]) {
        case 'MusicDataForSportsFes':
                try {
                    let class_name = db_sheet.getRange(findRow(db_sheet,arguments[3],5),2).getValue();
                    let classSheetName = class_name + ".指示情報(体育祭)";
                    let sheet = ss.getSheetByName(classSheetName);

                    sheet.getRange("A17").setValue(arguments[1]);
                    sheet.getRange("E17").setValue(arguments[2]);

                    for(let i = 0; i < arguments[4].length && i < arguments[5].length && i < arguments[6].length && i < arguments[7].length; i++) {
                        let nameRange = sheet.getRange(i + 8, 2);
                        let timeRange = sheet.getRange(i + 8, 6);
                        let fadeInRange = sheet.getRange(i + 8, 8);
                        let fadeOutRange = sheet.getRange(i + 8, 9);
                        nameRange.setValue(arguments[4][i]);
                        timeRange.setValue(arguments[5][i]);
                        fadeInRange.setValue(arguments[6][i]);
                        fadeOutRange.setValue(arguments[7][i]);
                    }

                    let msg = "HTTPステータス : 200 OK<br />送信されました。ページを閉じていただいて構いません。いつでも指示情報を変更することはできますが、期日は守ってください。<br /><br />"
                    return msg;
                } catch {
                    let msg = "HTTPステータス : 400 Bad Request<br />エラーが発生しました。もう一度お試しください。";
                    return msg;
                }

        case 'classData':
                let ifErrorSheetsDel = [];
                if(arguments[1].includes("体育祭")) {
                    try {
                        const baseSheet = ss.getSheetByName("S_F-db");
                        baseSheet.copyTo(ss).setName(arguments[2] + ".指示情報(体育祭)");

                        let classSheetName = arguments[2] + ".指示情報(体育祭)";
                        ss.getSheetByName(classSheetName).getRange("B4").setValue(arguments[2]);
                        ss.getSheetByName(classSheetName).getRange("G4").setValue(arguments[4] + " " + arguments[5]);
                        ss.getSheetByName(classSheetName).getRange("G5").setValue(arguments[3]);
                        db_sheet.appendRow([arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]]);

                        let msg = "HTTPステータス : 200 OK<br />"
                        return {msg, classSheetName};

                    } catch {
                        let classSheetName = "error";
                        let sheets = ss.getSheets();
                        for(let i = 0; i < sheets.length; i++) {
                            if(sheets[i].getName().includes("のコピー")) {
                                ifErrorSheetsDel.push(sheets[i])
                            }
                        }
                        for(let j = 0; j < ifErrorSheetsDel.length; j++) {
                            ss.deleteSheet(ifErrorSheetsDel[j]);
                        }
                        let msg = "HTTPステータス : 406 Not Acceptable<br />その団体名はすでに登録されています。";
                        return {msg, classSheetName};
                    }
                } else {
                    try {
                        const baseSheet = ss.getSheetByName("db");
                        baseSheet.copyTo(ss).setName(arguments[2] + ".指示情報");

                        let classSheetName = arguments[2] + ".指示情報";
                        let classData = ss.getSheetByName(classSheetName).getRange("A8:H27").getValues();
                        ss.getSheetByName(classSheetName).getRange("B4").setValue(arguments[2]);
                        ss.getSheetByName(classSheetName).getRange("G4").setValue(arguments[4] + " " + arguments[5]);
                        ss.getSheetByName(classSheetName).getRange("G5").setValue(arguments[3]);
                        db_sheet.appendRow([arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]]);

                        let msg = "HTTPステータス : 200 OK<br />"
                        return {msg, classData, classSheetName};
                    } catch {
                        let classData = "error";
                        let classSheetName = "error";
                        let sheets = ss.getSheets();
                        for(let i = 0; i < sheets.length; i++) {
                            if(sheets[i].getName().includes("のコピー")) {
                              ifErrorSheetsDel.push(sheets[i])
                            }
                        }
                        for(let j = 0; j < ifErrorSheetsDel.length; j++) {
                            ss.deleteSheet(ifErrorSheetsDel[j]);
                        }
                        let msg = "HTTPステータス : 409 Conflict<br />その団体名はすでに登録されています。"
                        return {msg, classData, classSheetName};
                    }
                }

        case 'AdminOperation':
                try {
                    adminSheet.getRange("C7").setValue(arguments[1]);
                    adminSheet.getRange("C8").setValue(arguments[2]);
                    db_sheet.getDataRange().setValue("");
                    db_sheet.appendRow(["行事名", "団体名", "アドレス", "学籍番号", "氏名"])

                    arguments[3].forEach(row => {
                        db_sheet.appendRow(row);
                    });

                    let msg = "HTTPステータス : 200 OK<br />";
                    return msg;
                } catch {
                    let msg = "HTTPステータス : 202 Accepted<br />";
                    return msg;
                }

        case 'addClass':
                try {
                    let eventName = adminSheet.getRange("C7").getValue();
                    db_sheet.appendRow([eventName, arguments[1], arguments[2], arguments[3], arguments[4]])
                    let msg = "HTTPステータス : 200 OK<br />";
                    return msg;
                } catch {
                    let msg = "HTTPステータス : 202 Accepted<br />"
                    return msg;
                }
            

        case 'Gmail':
                let header = `<strong>こんにちは、${arguments[3]}さん。</strong><br /><hr /><br /><br />`;
                let body = `<font size="4">${arguments[4]}</font><br />`;
                let footer = '<br /><br /><hr /><strong>放送部行事支援ツール Plform</strong><br />';
                let draft = GmailApp.createDraft(arguments[2], arguments[1], "body", {
                                        name: "Plform",
                                        htmlBody: (header + body + footer).replaceAll('\n', '<br />'),
                                        });
                draft.send();
                return;
    }
}
