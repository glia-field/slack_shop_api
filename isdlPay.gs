//This Library ID: MF69OVvcBvkymVokVsE1aHeaMJ5Q-zlzu

//Requirement Library
//slackApp: M3W5Ut3Q39AaIwLquryEPMwV62A3znfOO

function transMoney(recvId, sendId, value, slack_access_token, sheet_id) {
  //get user information in JSON.
  var sheet = SpreadsheetApp.openById(sheet_id);
  var recvInfo = getInfo(slack_access_token, recvId, sheet);
  var sendInfo = getInfo(slack_access_token, sendId, sheet);
  
  //どちらかが新規ユーザーであった場合はsheetを再読込
  if(recvInfo.newUserCheck || sendInfo.newUserCheck){
    sheet = SpreadsheetApp.openById(sheet_id);
  }
  
  //spreadSheetに増減後の値を入力
  sheet.getRange(recvInfo.sheetMoneyAddress).setValue(parseInt(recvInfo.money) + value);
  sheet.getRange(sendInfo.sheetMoneyAddress).setValue(parseInt(sendInfo.money) - value);
  
  postMessage(slack_access_token, "@"+recvId,"残高:"+(parseInt(recvInfo.money) + value)+"[+"+value+"]");
  postMessage(slack_access_token, "@"+sendId,"残高:"+(parseInt(sendInfo.money) - value)+"[-"+value+"]");
  postMessage(slack_access_token, "#money_log","[送金]"+sendInfo.userName+"->"+recvInfo.userName+"["+value+"円]");
}

function addMoney(userId, value, slack_access_token, sheet_id) {
  //get user information in JSON.
  var sheet = SpreadsheetApp.openById(sheet_id);
  var userInfo = getInfo(slack_access_token, userId, sheet);
  
  //新規ユーザーであった場合はsheetを再読込
  if(userInfo.newUserCheck){
    sheet = SpreadsheetApp.openById(sheet_id);
  }
  
  //spreadSheetに増額後の値を入力
  sheet.getRange(userInfo.sheetMoneyAddress).setValue(parseInt(userInfo.money) + value);
    
  postMessage(slack_access_token, "@"+userId,"残高:"+money+"[+"+value+"]");
  postMessage(slack_access_token, "#money_log","[入金]"+userInfo.userName+"残高:"+money+"[+"+value+"]");
}

function subMoney(userId, value, slack_access_token, sheet_id) {
  //get user information in JSON.
  var sheet = SpreadsheetApp.openById(sheet_id);
  var userInfo = getInfo(slack_access_token, userId, sheet);
  
  //新規ユーザーであった場合はsheetを再読込
  if(userInfo.newUserCheck){
    sheet = SpreadsheetApp.openById(sheet_id);
  }
  
  //spreadSheetに増額後の値を入力
  sheet.getRange(userInfo.sheetMoneyAddress).setValue(parseInt(userInfo.money) - value);
    
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

function getInfo(slack_access_token, userId, sheet){
  var app = SlackApp.create(slack_access_token); 
                                                       
  var lastrow = sheet.getLastRow();
  var userIdList = sheet.getSheetValues(1, 1, lastrow, 1);
  var moneyList = sheet.getSheetValues(1, 2, lastrow, 1);
  var userNameList = sheet.getSheetValues(1, 3, lastrow, 1);
  var indexNum = arrayParse(userIdList).indexOf(userId);
  
  //新規ユーザーの検出
  var newUserCheck = false; //新規ユーザーか否かのフラグ
  if(indexNum >= 0){
    var money = moneyList[indexNum];
    var userName = userNameList[indexNum];
  }else{
    //新規ユーザーを仮追加
    var lastrow = userIdList.length;
    sheet.getRange("A"+(lastrow+1)).setValue(userId);
    sheet.getRange("B"+(lastrow+1)).setValue("0");
    
    //SlackのWeb APIでユーザーIDから各種情報を取得
    var userInfo = app.usersInfo(userId);
    var userRealName = userInfo.user.profile.real_name;
    var userName = userInfo.user.name;
    
    sheet.getRange("C"+(lastrow+1)).setValue(userRealName);
    sheet.getRange("D"+(lastrow+1)).setValue("@"+userName);
    
    //例: Excelのn行目に新規ユーザーを追加した場合、lastrowには(n-1)が入っており、新規ユーザーは配列的には(n-1)に追加されたことにななる。
    indexNum = lastrow;
    
    var money = "0";
    var userName = userRealName;
    newUserCheck = true;
  }
  
  //JSONにして返す
  var info = {
    "userId": userId,
    "userName" : userName,
    "money": money,
    "indexNum": indexNum,
    "sheetMoneyAddress": "B"+(indexNum+1),
    "newUserCheck": newUserCheck
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