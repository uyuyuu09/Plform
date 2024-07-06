function findRow(sheet, val, col){
    var dat = sheet.getDataRange().getValues();
     
    for (var i = 1; i < dat.length; i++) {
        if (dat[i][col - 1] == val) {
            return i + 1;
        }
    }
    return 0;
}

function findMultiRow(sheet, val, col) {
    var dat = sheet.getDataRange().getValues();
    var targetRows = [];
    var data = [];
    for (var i = 0; i < dat.length; i++) {
        if (dat[i][col - 1] == val) {
            targetRows.push(i + 1);
        }
    }
    targetRows = Array.from(new Set(targetRows));
    for (let i = 0; i < targetRows.length; i++) {
        // 検索にヒットしたレコードの取得
        let tmpdata = sheet.getRange(targetRows[i], 1, 1, sheet.getLastColumn()).getValues();
        data.push(tmpdata[0]);
    }
    return data;
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
    // const KBC_DB = ss.getSheetByName("部内名簿");
    try {
        var user_name = member.getRange(findRow(member,loginUserGmail,3),2).getValue();
        var class_num = member.getRange(findRow(member,loginUserGmail,3),1).getValue();
        // var department = KBC_DB.getRange(findRow(KBC_DB,loginUserGmail,5),3).getValue();
    } catch {
        var user_name = "ゲストさん";
        var class_num = "9999"
        // var department = "外部";
    }
    
    const adminSheet = ss.getSheetByName("admin");
    const eventName = adminSheet.getRange("C7").getValue();
    const deadLine = adminSheet.getRange("C8").getValue();
    
    temp.page = parameter;
    temp.url = ScriptApp.getService().getUrl();
    temp.gmail = loginUserGmail;
    temp.user = user_name;
    temp.class_num = class_num;
    // temp.department = department;
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
                    // let mainMusicData = ss.getSheetByName(classSheetName).getRange("A20").getValue();
                    // let mainMusicTimeData = ss.getSheetByName(classSheetName).getRange("E20").getValue();
                    // let classMusicNameData = ss.getSheetByName(classSheetName).getRange("B8:B17").getValues();
                    // let classMusicTimeData = ss.getSheetByName(classSheetName).getRange("F8:F17").getValues();
                    // let musicFadeInData = ss.getSheetByName(classSheetName).getRange("H8:H17").getValues();
                    // let musicFadeOutData = ss.getSheetByName(classSheetName).getRange("I8:I17").getValues();
                    // let remarksForMusic = ss.getSheetByName(classSheetName).getRange("H20").getValue();

                    //After
                    let classSheet = ss.getSheetByName(classSheetName);
                    let data = classSheet.getDataRange().getValues();
                    let mainMusicData = data[19][0];
                    let mainMusicTimeData = data[19][4];
                    let remarksForMusic = data[19][7];
                    let classMusicNameData = classSheet.getRange("B8:B17").getValues();
                    let classMusicTimeData = classSheet.getRange("F8:F17").getValues();
                    let musicFadeInData = classSheet.getRange("H8:H17").getValues();
                    let musicFadeOutData = classSheet.getRange("I8:I17").getValues();

                    return {mainMusicData, mainMusicTimeData, classMusicNameData, classMusicTimeData, classSheetName, musicFadeInData, musicFadeOutData, remarksForMusic};
                } else {
                    let classSheetName = class_name + ".指示情報";
                    let classSheet = ss.getSheetByName(classSheetName);
                    let classData = classSheet.getRange("A8:H27").getValues();
                    let data = classSheet.getRange("A34:G60").getValues();
                    let mic_Cab = data[0][4];
                    let mic_WL = data[2][4];
                    let micStand_Mini = data[4][4];
                    let micStand_Big = data[6][4];
                    let spot_left = data[8][4];
                    let spot_right = data[10][4];
                    let light = data[12][2];
                    let projector = data[14][2];
                    let remarksForEvents = data[21][6];

                    let musicNameData = classSheet.getRange("B55:B57").getValues();
                    let musicTimeData = classSheet.getRange("D55:D57").getValues();

                    return {classData, classSheetName, mic_Cab, mic_WL, micStand_Mini, micStand_Big, spot_left, spot_right, light, projector, musicNameData, musicTimeData, remarksForEvents};
                }
        case 'AdminInfomation':
                let AdminLoginId = adminSheet.getRange("C9").getValue();
                
                if(arguments[1] === AdminLoginId) {
                    let data = adminSheet.getRange("C7:C8").getValues();
                    let eventName = data[0][0];
                    let deadLine = data[1][0];
                    let dbInfo = db_sheet.getDataRange().getValues().slice(1);
                    let names = ss.getSheets().map(function(sheet) {
                      return sheet.getName();
                    });
                    let sheetNames = names.join(",");
                    let msg = "HTTPステータス : 200 OK<br />"
                    return {msg, eventName, deadLine, dbInfo, sheetNames}
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

                    sheet.getRange("A20").setValue(arguments[1]);
                    sheet.getRange("E20").setValue(arguments[2]);
                    sheet.getRange("H20").setValue(arguments[8]);

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

        case 'dataForOtherEvents':
                try {
                    let class_name = db_sheet.getRange(findRow(db_sheet,arguments[1],5),2).getValue();
                    let classSheetName = class_name + ".指示情報";
                    let sheet = ss.getSheetByName(classSheetName);

                    if(arguments[2] === "" && arguments[3] === "") {
                        sheet.getRange("C34").setValue("なし");
                        sheet.getRange("C38").setValue("なし");
                        sheet.getRange("E34").setValue("なし");
                        sheet.getRange("E36").setValue("なし");
                        sheet.getRange("E38").setValue("なし");
                        sheet.getRange("E40").setValue("なし");
                    } else {
                        sheet.getRange("C34").setValue("あり");
                        sheet.getRange("C38").setValue("あり");
                        sheet.getRange("E34").setValue(arguments[2]);
                        sheet.getRange("E36").setValue(arguments[3]);
                        sheet.getRange("E38").setValue(arguments[4]);
                        sheet.getRange("E40").setValue(arguments[5]);
                    }

                    if(arguments[6].includes("使用")) {
                        sheet.getRange("C42").setValue("あり");
                        switch(arguments[6]) {
                            case "下手側・上手側使用":
                                    sheet.getRange("E42").setValue("あり");
                                    sheet.getRange("E44").setValue("あり");
                                    break;
                            case "下手側のみ使用":
                                    sheet.getRange("E42").setValue("あり");
                                    sheet.getRange("E44").setValue("なし");
                                    break;
                            case "上手側のみ使用":
                                    sheet.getRange("E42").setValue("なし");
                                    sheet.getRange("E44").setValue("あり");
                                    break;
                        }
                    } else {
                        sheet.getRange("C42").setValue("なし");
                        sheet.getRange("E42").setValue("なし");
                        sheet.getRange("E44").setValue("なし");
                    }
                    
                    sheet.getRange("C46").setValue(arguments[7]);
                    sheet.getRange("C48").setValue(arguments[8]);
                    sheet.getRange("G55").setValue(arguments[9]);

                    for(let i = 0; i < arguments[10].length && i < arguments[11].length; i++) {
                        let nameRange = sheet.getRange(i + 55, 2);
                        let timeRange = sheet.getRange(i + 55, 4);
                        nameRange.setValue(arguments[10][i]);
                        timeRange.setValue(arguments[11][i]);
                    }

                    let msg = "HTTPステータス : 200 OK<br />送信されました。ページを閉じていただいて構いません。いつでも指示情報を変更することはできますが、期日は守ってください。<br /><br />";
                    return msg;
                } catch {
                    let msg = "HTTPステータス : 400 Bad Request<br />エラーが発生しました。もう一度お試しください。";
                    return msg;
                }

        case 'classData':
                let ifErrorSheetsDel = [];
                judge = findMultiRow(db_sheet, arguments[5], 5);
                if(arguments[1].includes("体育祭")) {
                    try {
                        if(judge.length === 0) {
                            const baseSheet = ss.getSheetByName("S_F-db");
                            baseSheet.copyTo(ss).setName(arguments[2] + ".指示情報(体育祭)");

                            let classSheetName = arguments[2] + ".指示情報(体育祭)";
                            ss.getSheetByName(classSheetName).getRange("B4").setValue(arguments[2]);
                            ss.getSheetByName(classSheetName).getRange("G4").setValue(arguments[4] + " " + arguments[5]);
                            ss.getSheetByName(classSheetName).getRange("G5").setValue(arguments[3]);
                            db_sheet.appendRow([arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]]);

                            let msg = "HTTPステータス : 200 OK<br />"
                            return {msg, classSheetName};
                        } else {
                            let classSheetName = "error";
                            let msg = "HTTPステータス : 406 Not Acceptable<br />あなたの氏名はすでに登録されています。";
                            return {msg, classSheetName};
                        }
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
                        if(judge.length === 0) {
                            const baseSheet = ss.getSheetByName("temp");
                            baseSheet.copyTo(ss).setName(arguments[2] + ".指示情報");

                            let classSheetName = arguments[2] + ".指示情報";
                            let classData = ss.getSheetByName(classSheetName).getRange("A8:H27").getValues();
                            ss.getSheetByName(classSheetName).getRange("B4").setValue(arguments[2]);
                            ss.getSheetByName(classSheetName).getRange("G4").setValue(arguments[4] + " " + arguments[5]);
                            ss.getSheetByName(classSheetName).getRange("G5").setValue(arguments[3]);
                            db_sheet.appendRow([arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]]);

                            let msg = "HTTPステータス : 200 OK<br />"
                            return {msg, classData, classSheetName};
                        } else {
                            let classSheetName = "error";
                            let msg = "HTTPステータス : 406 Not Acceptable<br />あなたの氏名はすでに登録されています。";
                            return {msg, classSheetName};
                        }
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
                    let msg = "HTTPステータス : 202 Accepted<br />送信されましたが、エラーが発生しました。<br />";
                    return msg;
                }

        case 'addClass':
                judge = findMultiRow(db_sheet, arguments[1], 2)
                try {
                    if(judge.length === 0) {
                        let eventName = adminSheet.getRange("C7").getValue();
                        db_sheet.appendRow([eventName, arguments[1], arguments[2], arguments[3], arguments[4]])
                        let msg = "HTTPステータス : 200 OK<br />";
                        return msg;
                    } else {
                        let msg = "HTTPステータス : 406 Not Acceptable<br />その団体名はすでに登録されています。";
                        return msg;
                    }
                } catch {
                    let msg = "HTTPステータス : 202 Accepted<br />送信されましたが、エラーが発生しました。<br />";
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

function deleteDoneSheets() {;
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheets = ss.getSheets();
    let deleteSheets = [];

    for(let i = 0; i < sheets.length; i++) {
        if(sheets[i].getName().includes("指示情報")) {
            deleteSheets.push(sheets[i])
        } else {
            //指示情報シート以外は除外
        }
    }

    for(let j = 0; j < deleteSheets.length; j++) {
        ss.deleteSheet(deleteSheets[j])
    }

    let sheet = ss.getSheetByName("団体情報記録");
    sheet.getDataRange().setValue("");
    sheet.appendRow(["行事名", "団体名", "アドレス",	"学籍番号",	"氏名"])
}
