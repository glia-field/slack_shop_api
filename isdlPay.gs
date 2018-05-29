//This Library ID: MF69OVvcBvkymVokVsE1aHeaMJ5Q-zlzu

//Requirement Library
//slackApp: M3W5Ut3Q39AaIwLquryEPMwV62A3znfOO
function transMoney(sendUserId, recvUserId, value, slack_access_token, sheet_id) {
    var app = SlackApp.create(slack_access_token);
    
    // マスタデータシートを取得
    var sheet = SpreadsheetApp.openById(sheet_id);　
    var lastrow = sheet.getLastRow();
    var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
    var memberName = sheet.getSheetValues(1, 3, lastrow, 1); //データ行のみを取得する
    　　　　　　
    var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
    var sendUserIndexNum = arrayParse(member).indexOf(sendUserId);
    var recvUserIndexNum = arrayParse(member).indexOf(recvUserId);
    var sendUserMoney = parseInt(money[sendUserIndexNum]) - value;
    var recvUserMoney = parseInt(money[recvUserIndexNum]) + value;
    sheet.getRange(sendUserIndexNum + 1, 2).setValue(sendUserMoney);
    sheet.getRange(recvUserIndexNum + 1, 2).setValue(recvUserMoney);
    postMessage(slack_access_token, app, "@" + sendUserId, "残高:" + sendUserMoney + "[-" + value + "]");
    postMessage(slack_access_token, app, "@" + recvUserId, "残高:" + recvUserMoney + "[+" + value + "]");
    postMessage(slack_access_token, app, "#money_log", "[送金] @" + memberName[sendUserIndexNum] + "-> @" + memberName[recvUserIndexNum] + "[" + value + "円]");
}

function addMoney(userId, value, slack_access_token, sheet_id) {
    var app = SlackApp.create(slack_access_token);
  
    //spreadsheetの読み込み
    var sheet = SpreadsheetApp.openById(sheet_id);
    var lastrow = sheet.getLastRow();
    var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
    
    //残高の列を取得し、今回のユーザーが何行目かを探す
    var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
    var indexNum = arrayParse(member).indexOf(userId);

    //該当者の残高をint型数字に整形して、今回受け取った額を足し合わせ、スプレッドシートに記述
    money = parseInt(money[indexNum]) + value;
    sheet.getRange(indexNum + 1, 2).setValue(money);
    
    //通知
    postMessage(slack_access_token, app, "@" + userId, "残高:" + money + "[+" + value + "]");
    postMessage(slack_access_token, app, "#money_log", "[入金]" + getNameById(userId, sheet_id) + "残高:" + money + "[+" + value + "]");
}

function subMoney(userId, value, slack_access_token, sheet_id) {
    var app = SlackApp.create(slack_access_token);
  
    //spreadsheetの読み込み
    var sheet = SpreadsheetApp.openById(sheet_id);
    var lastrow = sheet.getLastRow();
    var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
    
    //残高の列を取得し、今回のユーザーが何行目かを探す
    var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
    var indexNum = arrayParse(member).indexOf(userId);

    //該当者の残高をint型数字に整形して、今回受け取った額を足し合わせ、スプレッドシートに記述
    money = parseInt(money[indexNum]) - value;
    sheet.getRange(indexNum + 1, 2).setValue(money);
    
    //通知
    postMessage(slack_access_token, app, "@" + userId, "残高:" + money + "[-" + value + "]");
    postMessage(slack_access_token, app, "#money_log", "[入金]" + getNameById(userId, sheet_id) + "残高:" + money + "[-" + value + "]");
}

function getMoney(userId, sheet_id) {
  //spreadsheetの読み込み
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
  var money_old = money.slice(0, money.length);
  var indexNum = arrayParse(member).indexOf(userId);
  return money[indexNum][0];
}

function getIdByName(userName, sheet_id) {
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  var member_name = sheet.getSheetValues(1, 3, lastrow, 1);
  var indexNum = arrayParse(member_name).indexOf(userName);
  return member[indexNum][0];
}

function getNameById(userId, sheet_id) {
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  var member_name = sheet.getSheetValues(1, 3, lastrow, 1);
  var indexNum = arrayParse(member).indexOf(userId);
  return member_name[indexNum][0];
}

function getIdByAtname(atname, sheet_id) {
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  var member_atname = sheet.getSheetValues(1, 4, lastrow, 1);
  var indexNum = arrayParse(member_atname).indexOf(atname);
  return member[indexNum][0];
}

function getMemberList(sheet_id) {
  // マスタデータシートを取t得
  　　　
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  return member;
}

function getMoneyList(sheet_id) {
  // マスタデータシートを取t得
  　　
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
  return money;
}

function arrayParse(array) {
  var parseArray = [];
  for (var i = 0; i < array.length; i++) {
    parseArray[i] = array[i][0];
  }
  return parseArray;
}

function postMessage(token, app, id, message) {
  var bot_name = "ウィーゴ";
  var bot_icon = "http://www.hasegawa-model.co.jp/hsite/wp-content/uploads/2016/04/cw12p5.jpg";
  return app.postMessage(id, message, {
    username: bot_name,
    icon_url: bot_icon,
    link_names: 1
  });
}