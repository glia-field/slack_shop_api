function addMoney(userId,value) {
  
  if(value > 0){  
    var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
    var app = SlackApp.create(token);  
    
    //spreadsheetの読み込み
    var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    var sheet = SpreadsheetApp.openById(sheet_id);
    var lastrow = sheet.getLastRow();
    var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
    var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
    
    var indexNum = arrayParse(member).indexOf(userId);
    money = parseInt(money[indexNum]) + value;
    
    var Address = "B"+(indexNum+1);
    sheet.getRange(Address).setValue(money);
    
    postMessage(token,app, "@"+userId,"残高:"+money+"[+"+value+"]");
    postMessage(token,app, "#money_log","[入金]"+getNameById(userId)+"残高:"+money+"[+"+value+"]");
  }
}

function subMoney(userId,value) {
  if(value > 0){
    var token = PropertiesService.getScriptProperties().getProperty('SLACK_ACCESS_TOKEN');
    var app = SlackApp.create(token);  
    
    //spreadsheetの読み込み
    var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    var sheet = SpreadsheetApp.openById(sheet_id);
    var lastrow = sheet.getLastRow();
    var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
    var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
  
    var indexNum = arrayParse(member).indexOf(userId);
    money = parseInt(money[indexNum]) - value;
  
    var Address = "B"+(indexNum+1);
    sheet.getRange(Address).setValue(money);
  
    postMessage(token,app, "@"+userId,"残高:"+money+"[-"+value+"]");
    if(money < 0){
      postMessage(token,app, "@"+userId,"残高がありません。チャージしてくださいね");
    }
    postMessage(token,app, "#money_log","[出金]"+getNameById(userId)+"残高:"+money+"[-"+value+"]");
  }
}

function getMoney(userId) {
  //spreadsheetの読み込み
  var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
  var money = sheet.getSheetValues(1, 2, lastrow, 1); //データ行のみを取得する
  var money_old = money.slice(0, money.length);
  
  var indexNum = arrayParse(member).indexOf(userId);
  return money[indexNum][0];
}

function getIdByName(userName){
  var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
  var member_name = sheet.getSheetValues(1,3,lastrow,1);
  
  var indexNum = arrayParse(member_name).indexOf(userName);
  return member[indexNum][0];
}

function getNameById(userId){
  var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
  var member_name = sheet.getSheetValues(1,3,lastrow,1);
  
  var indexNum = arrayParse(member).indexOf(userId);
  return member_name[indexNum][0];
}

function getIdByAtname(atname){
  var sheet_id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  var sheet = SpreadsheetApp.openById(sheet_id);
  var lastrow = sheet.getLastRow();
  var member = sheet.getSheetValues(1, 1, lastrow, 1);  //データ行のみを取得する
  var member_atname = sheet.getSheetValues(1,4,lastrow,1);
  
  var indexNum = arrayParse(member_atname).indexOf(atname);
  return member[indexNum][0];
}

function arrayParse(array){
  var parseArray = [];
  for(var i=0; i<array.length; i++){
    parseArray[i] = array[i][0]; 
  }
  
  return parseArray;
}

function postMessage(token,app, id,message){
  var bot_name = "ウィーゴ";
  var bot_icon = "http://www.hasegawa-model.co.jp/hsite/wp-content/uploads/2016/04/cw12p5.jpg"; 
  
  return app.postMessage(id, message, {
    username: bot_name,
    icon_url: bot_icon,
    link_names: 1
  });
}