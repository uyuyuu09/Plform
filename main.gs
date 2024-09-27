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

  var member = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("名簿");
  var LOGIN_USER = Session.getActiveUser().getEmail();
  try {
    var user_name = member.getRange(findRow(member,LOGIN_USER,3),2).getValue();
    var department = member.getRange(findRow(member,LOGIN_USER,3),5).getValue();
  } catch {
    var user_name = "unknown";
    var department = "unknown";
    const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAiRyHcdc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PUM49NtMjOvYVEfRc1zFLeusUyeycrSKvQ7RWyS3r9s";
  
    let msg = ["放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。\nデータベースに登録されていない人からのログイン試行を検知しました。"];
    var message = {
      'text' : msg.join('\n')
    };

    var options = {
      'payload' : JSON.stringify(message),
      'method' : 'POST',
      'contentType' : 'application/json'
    };

    var response = UrlFetchApp.fetch(WEBHOOK_URL,options);
  }

  if(LOGIN_USER.includes("220109")) {
    department = 'admin';
  }

  var shift = member.getRange(findRow(member,LOGIN_USER,3),6).getValue();

  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var names = spreadsheet.getSheets().map(function(sheet) {
    return sheet.getName();
  });
  let sheetNames = names.join(",");

  temp.page = parameter;
  temp.user_name = user_name;
  temp.department = department;
  temp.shift = shift;
  temp.sheetNames = sheetNames;
  temp.url = ScriptApp.getService().getUrl();
  let res = temp
                .evaluate()
                .setTitle('SIRIUS')
                .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=1.0');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('アクセス履歴');
  const nowTime = new Date();
  const [month,
         date,
         youbi,
         hour,
         minute,
         second
        ] = [nowTime.getMonth() + 1,
             nowTime.getDate(),
             ["日", "月", "火", "水", "木", "金", "土"][nowTime.getDay()],
             nowTime.getHours(),
             nowTime.getMinutes(),
             nowTime.getSeconds()
            ];
  sheet.appendRow([sheet.getLastRow(), LOGIN_USER, user_name, department, month, date, youbi, hour, minute, second, temp.url]);

  return res;
} 

function set2fig(num) {
  return String(num).padStart(2, '0');
}

function generateQRCode(url) {
  var imageUrl = "https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=" + encodeURIComponent(url);
  return imageUrl;
}

function sendLog(msg) {
  if (msg === "") {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('log');
  const member = ss.getSheetByName("名簿");
  const LOGIN_USER = Session.getActiveUser().getEmail();
  const user_name = member.getRange(findRow(member, LOGIN_USER, 3), 2).getValue();
  let department = member.getRange(findRow(member, LOGIN_USER, 3), 5).getValue();

  if (LOGIN_USER.includes("220109")) {
    department = '管理者';
  }

  const nowTime = new Date();
  const [month, date, youbi, hour, minute] = [
    nowTime.getMonth() + 1,
    nowTime.getDate(),
    ["日", "月", "火", "水", "木", "金", "土"][nowTime.getDay()],
    String(nowTime.getHours()).padStart(2, "0"),
    String(nowTime.getMinutes()).padStart(2, "0"),
  ];

  sheet.appendRow([sheet.getLastRow() + 1, user_name, department, msg, " ", month + "/" + date + "(" + youbi + ")" + "   " + hour + ":"  + minute]);
}


function sendAdminOperation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("部署管理");
  let data = sheet.getDataRange().getValues(); 

  for(let i = 0; i < data.length; i++) {
    if(data[i][1] === "ready") {
      sheet.getRange(i + 1, 2).setValue("not ready");
      console.log("success");
    } else {
      console.log("error");
    }
  }
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mem_DB = ss.getSheetByName("名簿");
  let log_DB = ss.getSheetByName("log");
  let vertically_DB = ss.getSheetByName("団体詳細一覧");
  let kitakou_collection_DB = ss.getSheetByName("北高コレクション");
  let kita_colle_book_DB = ss.getSheetByName("北コレ予約");
  let LOGIN_USER = Session.getActiveUser().getEmail();
  switch (arguments[0]) {
    case 'user_name':
      user_name = mem_DB.getRange(findRow(mem_DB, LOGIN_USER, 3), 2).getValue();
      return user_name;

    case 'shift':
      shift = mem_DB.getRange(findRow(mem_DB,LOGIN_USER,3),6).getValue();
      return shift;

    case 'mem_DB':
      mem = mem_DB.getRange("A1:F35").getValues();
      return mem;

    case 'log_DB':
      log = log_DB.getDataRange().getValues().reverse();
      return log;

    case 'vertically_DB':
      vertically_db = vertically_DB.getRange("A1:G9").getValues();
      return vertically_db;

    case 'kitakou_collection_DB':
      kitakou_collection_db = kitakou_collection_DB.getRange("A1:I98").getValues().slice(4);
      return kitakou_collection_db;

    case 'kitakou_collection_now':
      kitakou_collection_db = kitakou_collection_DB.getRange("A1:I98").getValues().slice(4);
      // カウント用の変数を初期化
      var countDone = 0;
      var countUnscheduled = 0;
      var countWaiting = 0;

      for (var i = 0; i < kitakou_collection_db.length; i++) {
        var status = kitakou_collection_db[i][4];

        if (status === '撮影済') {
          countDone++;
        } else if (status === 'アポ未取得') {
          countUnscheduled++;
        } else if (status === '撮影待ち') {
          countWaiting++;
        }
      }
      return {
        done: countDone,
        unscheduled: countUnscheduled,
        waiting: countWaiting
      };

    case 'kita_colle_book_DB':
      kita_colle_book_db = kita_colle_book_DB.getDataRange().getValues();
      return kita_colle_book_db;

    case 'all':
      const allData = {
        user_name: mem_DB.getRange(findRow(mem_DB, LOGIN_USER, 3), 2).getValue(),
        shift: mem_DB.getRange(findRow(mem_DB, LOGIN_USER, 3), 6).getValue(),
        mem_DB: mem_DB.getRange("A1:F35").getValues(),
        log_DB: log_DB.getDataRange().getValues(),
        vertically_DB: vertically_DB.getRange("A1:G9").getValues().slice(1),
        kitakou_collection_DB: kitakou_collection_DB.getRange("A1:I98").getValues().slice(4),
        kita_colle_book_DB: kita_colle_book_DB.getDataRange().getValues(),
        kitakou_collection_now: (() => {
          // カウント用の変数を初期化
          var countDone = 0;
          var countUnscheduled = 0;
          var countWaiting = 0;
          const kitakou_collection_db = kitakou_collection_DB.getRange("A1:I98").getValues().slice(4);

          for (var i = 0; i < kitakou_collection_db.length; i++) {
            var status = kitakou_collection_db[i][4];

            if (status === '撮影済') {
              countDone++;
            } else if (status === 'アポ未取得') {
              countUnscheduled++;
            } else if (status === '撮影待ち') {
              countWaiting++;
            }
          }
          return {
            done: countDone,
            unscheduled: countUnscheduled,
            waiting: countWaiting
          };
        })()
      };
    return allData;
  }
}

function updateKitakoreBook(row, col, newValue) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("北コレ予約");
  sheet.getRange(row + 1, col + 1).setValue(newValue);
}

function getAllSheetNames() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var names = spreadsheet.getSheets().map(function(sheet) {
    return sheet.getName();
  });
  let sheetNames = names.join(",");
  console.log(sheetNames)
  return sheetNames;
}

function savePdf(sheetName) {
  let folderId = "1tG8y3omcS_Txk3jZRYAco9WN9KgxgYb4"; //KBC大宮北高校放送部>01.行事運営>11.2023予餞会>test
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  let shId = ss.getSheetId();
  let fileName = sheetName;

  createPdf(folderId, ssId, shId, fileName);
}

function createPdf(folderId, ssId, shId, fileName){
  let baseUrl = "https://docs.google.com/spreadsheets/d/"
          + ssId
          + "/export?gid="
          + shId;
 
  let pdfOptions = "&exportFormat=pdf&format=pdf"
              + "&size=A4" //用紙サイズ (A4)
              + "&portrait=true"  //用紙の向き true: 縦向き / false: 横向き
              + "&fitw=true"  //ページ幅を用紙にフィットさせるか true: フィットさせる / false: 原寸大
              + "&top_margin=0.50" //上の余白
              + "&right_margin=0.50" //右の余白
              + "&bottom_margin=0.50" //下の余白
              + "&left_margin=0.50" //左の余白
              + "&horizontal_alignment=CENTER" //水平方向の位置
              + "&vertical_alignment=TOP" //垂直方向の位置
              + "&printtitle=false" //スプレッドシート名の表示有無
              + "&sheetnames=false" //シート名の表示有無
              + "&gridlines=false" //グリッドラインの表示有無
              + "&fzr=True" //固定行の表示有無
              + "&fzc=True" //固定列の表示有無
              ;

  let url = baseUrl + pdfOptions;

  let token = ScriptApp.getOAuthToken();

  let options = {
    headers: {
        'Authorization': 'Bearer ' +  token
    }
  };
 
  let blob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName + '.pdf');

  let folder = DriveApp.getFolderById(folderId);

  folder.createFile(blob);
}

function updateKitakoreSheet(row, col, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("北高コレクション");
  sheet.getRange(row + 4, col + 1).setValue(newValue);
}

function setKitakoreBook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kitakoreSheet = ss.getSheetByName("北高コレクション");
  const kitakoreBookSheet = ss.getSheetByName("北コレ予約");
  let sheetData = kitakoreSheet.getDataRange().getValues();
  let bookData = kitakoreBookSheet.getDataRange().getValues();
 
  for (let i = 0; i < bookData.length; i++) {
    for (let j = 1; j < bookData[i].length; j++) {
      let dataArray = bookData[i][j].split(',');
     
      for (let k = 0; k < dataArray.length; k++) {
        for (let l = 0; l < sheetData.length; l++) {
          if (sheetData[l][0] === dataArray[k]) {
            let dateCell = kitakoreBookSheet.getRange(i + 1, 1);
            let topCell = kitakoreBookSheet.getRange(1, j + 1);

            kitakoreSheet.getRange(l + 1, 8).setValue(dateCell.getValue() + '\n' + topCell.getValue());
          }
        }
      }
    }
  }
}

function googleChatBotTommorow() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAiRyHcdc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PUM49NtMjOvYVEfRc1zFLeusUyeycrSKvQ7RWyS3r9s";
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("北高コレクション");
 
  const today = new Date();
  const month = today.getMonth() + 1;

  let generateMonth = () => {
  if(month < 10) {
    return generatedMonth = "0" + month;
  } else {
    return generatedMonth = month;
  }};

  const day = today.getDate() + 1;
  
  let generateDay = () => {
  if(day < 10) {
    return generatedDay = "0" + day;
  } else {
    return generatedDay = day;
  }};
  
  const rows = ss.getRange('H1:H' + ss.getLastRow()).getValues();
  let msg = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0].includes(`${generateMonth()}/${generateDay()}`)) {
      const gData = ss.getRange(i + 1, 7).getValue();
      const bData = ss.getRange(i + 1, 2).getValue();
      const hData = ss.getRange(i + 1, 8).getValue();
      msg.push(gData + ' : ' + bData + '先生' + '  ' + '\n' + '<時間>' + '\n' + hData + '\n');
    }
  }
  msg = msg.length === 0 ? ["放送部ウェブシステム「SIRIUS」よりお知らせします。明日の撮影予定はありません。"] : [`放送部ウェブシステム「SIRIUS」よりお知らせします。明日の撮影は、\n\n${msg.join('\n')}\nです。`];
  var message = {
    'text' : msg.join('\n')
  };

  var options = {
    'payload' : JSON.stringify(message),
    'method' : 'POST',
    'contentType' : 'application/json'
  };
  var response = UrlFetchApp.fetch(WEBHOOK_URL,options);
}

function googleChatBotToday() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAiRyHcdc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PUM49NtMjOvYVEfRc1zFLeusUyeycrSKvQ7RWyS3r9s";
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("北高コレクション");
  const rows = ss.getRange('H1:H' + ss.getLastRow()).getValues();

  const today = new Date();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");

  let msg = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0].includes(`${month}/${day}`)) {
      const data = ss.getRange(`A${i + 1}:I${i + 1}`).getValues();
      msg.push(data[0][6] + ' : ' + data[0][1] + '先生' + '  ' + '\n' + '<時間>' + '\n' + data[0][7] + '\n');
    }
  }
  msg = msg.length === 0 ? ["放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。本日の撮影予定はありません。"] : [`放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。本日の撮影は、\n\n${msg.join('\n')}\nです。`];
  let message = {
    'text' : msg.join('\n')
  };
  let options = {
    'payload' : JSON.stringify(message),
    'method' : 'POST',
    'contentType' : 'application/json'
  };
  var response = UrlFetchApp.fetch(WEBHOOK_URL,options);
}

function googleChatBotUpdate() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAiRyHcdc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PUM49NtMjOvYVEfRc1zFLeusUyeycrSKvQ7RWyS3r9s";
 
  let msg = ["放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。\nSIRIUSがアップデートされました。ページを再読み込みしてください。"];
  var message = {
    'text' : msg.join('\n')
  };

  var options = {
    'payload' : JSON.stringify(message),
    'method' : 'POST',
    'contentType' : 'application/json'
  };

  var response = UrlFetchApp.fetch(WEBHOOK_URL,options);
}

function googleChatBotToday() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAiRyHcdc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PUM49NtMjOvYVEfRc1zFLeusUyeycrSKvQ7RWyS3r9s";
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("北高コレクション");
  const rows = ss.getRange('H1:H' + ss.getLastRow()).getValues();

  const today = new Date();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");

  let msg = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0].includes(`${month}/${day}`)) {
      const data = ss.getRange(`A${i + 1}:I${i + 1}`).getValues();
      msg.push(data[0][6] + ' : ' + data[0][1] + '先生' + '  ' + '\n' + '<時間>' + '\n' + data[0][7] + '\n');
    }
  }
  msg = msg.length === 0 ? ["放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。本日の撮影予定はありません。"] : [`放送部予餞会専用ウェブシステム「SIRIUS」よりお知らせします。本日の撮影は、\n\n${msg.join('\n')}\nです。`];
  let message = {
    'text' : msg.join('\n')
  };
  let options = {
    'payload' : JSON.stringify(message),
    'method' : 'POST',
    'contentType' : 'application/json'
  };
  var response = UrlFetchApp.fetch(WEBHOOK_URL,options);
}
