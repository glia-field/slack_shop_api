//This Library ID: MF69OVvcBvkymVokVsE1aHeaMJ5Q-zlzu

//Requirement Library
//slackApp: M3W5Ut3Q39AaIwLquryEPMwV62A3znfOO

function addMoney(userId, value, slack_access_token, sheet_id) {
  //get user information in JSON.
  var userInfo = getInfo(slack_access_token, userId, sheet_id);
    
  var moneyAddress = "B"+(userInfo.indexNum+1);
  var money = parseInt(userInfo.money) + value;
  var sheet = SpreadsheetApp.openById(sheet_id);
  sheet.getRange(moneyAddress).setValue(money);
    
  postMessage(slack_access_token, "@"+userId,"残高:"+money+"[+"+value+"]");
  postMessage(slack_access_token, "#money_log","[入金]"+userInfo.userName+"残高:"+money+"[+"+value+"]");
}

function subMoney(userId, value, slack_access_token, sheet_id) {
  //get user information in JSON.
  var userInfo = getInfo(slack_access_token, userId, sheet_id);
    
  var moneyAddress = "B"+(userInfo.indexNum+1);
  var money = parseInt(userInfo.money) - value;
  var sheet = SpreadsheetApp.openById(sheet_id);
  sheet.getRange(moneyAddress).setValue(money);
    
  postMessage(slack_access_token, "@"+userId,"残高:"+money+"[-"+value+"]");
  if(money < 0){
    postMessage(slack_access_token, "@"+userId,"残高がマイナスです。本システムは融資ではありません。");
  }
  postMessage(slack_access_token, "#money_log","[出金]"+userInfo.userName+"残高:"+money+"[-"+value+"]");
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
　var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1); //データ行のみを取得する
  return member;
}

function getMoneyList(sheet_id) {　
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

function getInfo(slack_access_token, userId, sheet_id){
  var app = SlackApp.create(slack_access_token); 
  var sheet = SpreadsheetApp.openById(sheet_id);
                                                                                                       
  var lastrow = sheet.getLastRow();
  var userIdList = sheet.getSheetValues(1, 1, lastrow, 1);
  var moneyList = sheet.getSheetValues(1, 2, lastrow, 1);
  var userNameList = sheet.getSheetValues(1, 3, lastrow, 1);
  var indexNum = arrayParse(userIdList).indexOf(userId);
  
  //新規ユーザーの検出
  if(indexNum==-1){
    //新規ユーザーを仮追加
    var lastrow = userIdList.length;
    sheet.getRange("A"+(lastrow+1)).setValue(userId);
    sheet.getRange("B"+(lastrow+1)).setValue("0");
    
    var userInfo = app.usersInfo(userId);
    var userName = userInfo.user.name;
    sheet.getRange("C"+(lastrow+1)).setValue(userName);
    sheet.getRange("D"+(lastrow+1)).setValue("@"+userName);
    
    //例: Excelのn行目に新規ユーザーを追加した場合、lastrowには(n-1)が入っており、新規ユーザーは配列的には(n-1)に追加されたことにななる。
    indexNum = lastrow;
    
    sheet = SpreadsheetApp.openById(sheet_id);
    lastrow = sheet.getLastRow();
    userIdList = sheet.getSheetValues(1, 1, lastrow, 1);
    moneyList = sheet.getSheetValues(1, 2, lastrow, 1);
    userNameList = sheet.getSheetValues(1, 3, lastrow, 1);
  }
  
  var info = {
    "userId": userId,
    "userName" : userNameList[indexNum],
    "money": moneyList[indexNum],
    "indexNum": indexNum
  }     
  
  return info;
}

function postMessage(slack_access_token, id, message) {
  var app = SlackApp.create(slack_access_token);  
  var bot_name = "ウィーゴ";
  var bot_icon = "http://www.hasegawa-model.co.jp/hsite/wp-content/uploads/2016/04/cw12p5.jpg";
  return app.postMessage(id, message, {
    username: bot_name,
    icon_url: bot_icon,
    link_names: 1
  });
}