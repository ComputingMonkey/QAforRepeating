///トークンを入力
var token = 'lpN1xhZhTsavqvb2qpw9LA07eAwsoNJ1pEFWQQLhCqn';//中野民
var token = 'pvSRCJomAef1Z1fLgtHidNoyOcxJ5yooAKE2i7mIqnz'//非本質LINE
var token = 'Cs0OtiBx92xR8DPTVuuXzjXTpBt9kXoVr0RbWqlsYfF';//2021非本質LINE

function test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var token = 'lpN1xhZhTsavqvb2qpw9LA07eAwsoNJ1pEFWQQLhCqn';
  var  message ='%';
    var options =
   {
     "method"  : "post",
     "payload" : "message=" + message,
     "headers" : {"Authorization" : "Bearer "+ token}
     
   };
  //Logger.log(message);
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}

//#回答率を算出
function getPercent(){
  //MJメンバーのデータリストを取得
  var sheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  //分母を定義（会員数）
  var deno = sheet.getLastRow();//47
  Logger.log(deno);
  //分子を定義（未回答者数）
  var nume = getYetMails().length;
  Logger.log(nume);
  //未回答者数/会員数の余事象の割合
  var percent_f = (1 - (nume / deno)) * 100;
  var percent = Math.round(percent_f * 1000)/1000;
  Logger.log(percent);
  return percent;
}

//#アンケートのメアドの列を取得
function getMailCol() {
  var sh = SpreadsheetApp.getActiveSheet();
  //検索するカテゴリ
  var key = 'メールアドレス';
  //タイトルの配列を取得
  var titles = sh.getRange(1,1,1,sh.getLastColumn()).getValues();
  //タイトルがある列を検索
  var key_col = titles[0].indexOf(key) + 1;
  //列の番号を返す
  return key_col;
}

//MJのメアドリストを下のリンクから引っ張ってくる
//https://docs.google.com/spreadsheets/d/1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw/edit#gid=0
function getAllMails(){
  var mailSheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
  //メアド配列を取得
  var allMails2d = mailSheet.getRange(1,1,mailSheet.getLastRow()).getValues();
  //二次元配列を一次元化
  var allMails = Array.prototype.concat.apply([],allMails2d);
  return allMails;
}


//#メアドと名前のhashを取得
function getHash(){   
    var sheet = SpreadsheetApp.openById('1GqTU4fCCcXn2Rlzfwj5oPxPG0Pzz1KPFP_WeDxMmcTw').getActiveSheet();
    var lastRow = sheet.getLastRow();
    // 配列の初期化
    var hashColor2 = {};
     
      for(var i = 2; i <= lastRow; i++) 
      {
            if(true)
            {  
              // 配列のkeyに対し値を設定する
              hashColor2[sheet.getRange(i, 1).getValue()] = sheet.getRange(i, 2).getValue();
            }
      }
  
        // 配列の要素(keyと値)を表示する
        for (var key in hashColor2) 
        {
            //Logger.log(key + "の値：" + hashColor2[key]);
        }
    //Logger.log(hashColor2);
    return hashColor2;
}


//#未回答メールの取得
function getYetMails() {
  //シート取得
  var mySheet = SpreadsheetApp.getActiveSheet();

  //#メールリスト呼び出し
  var allMails = getAllMails();
  
  //テスト用のメアド
  var testMails = ["allen.ackroyd1020@gmail.com","yuya.asano@aiesec.jp"];
  
  //#メアドの列を呼び出し
  var key_col = getMailCol();
  
  //回答済みメアド二次元配列を"自動"取得
  var doneMails2d = mySheet.getRange(2,key_col,mySheet.getLastRow() -1).getValues();
  
  //回答済み二次元配列→一次元配列に変換
  var doneMails1d = Array.prototype.concat.apply([],doneMails2d);
  
  //全メールリストと回答済み配列の差分を算出し、未回答者の配列を生成
  var yetMails = allMails.filter(function(e){return doneMails1d.filter(function(f){return e.toString() == f.toString()}).length == 0}); 
  
  return yetMails;
}

function getYetNames(){
  var hash = getHash();
  var yetMails = getYetMails();
  var yetNames = [];
  for(var i = 0;i < yetMails.length; i++){
    yetNames.push(hash[yetMails[i]]);
  }
  //Logger.log(yetNames.length);
  return yetNames;
}
  
//未回答者のメッセージボックスを表示
function putYetNames(){
  var yetNames = getYetNames();
  var message = '未回答者は以下の通りです\\n';
  Logger.log(yetNames.length);
  //未回答者のメッセージボックスを表示
  for(var i = 0;i < yetNames.length;i++){
    message += yetNames[i] + '\\n';
  }
  Logger.log(message);
  Browser.msgBox(message);
} 
//LineでNameを通知
function lineNotify(){

  //#アンケートのURLを呼び出し
  var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  //#未回答者リストの呼び出し
  var yetNames = getYetNames();
  //リストをLine用に変換
  var yetNames_m = '';
  for(var i = 0;i < yetNames.length;i++){
    var yetNames_m = yetNames_m + yetNames[i] + '\n';
  }
  //回答率を取得
  var percent = getPercent() + '％';
  //#送るメッセージを呼び出し
  var message = 
  '\n' + formUrl + '\n'
  + '【回答率:' +  percent + '】\n'
  + yetNames_m;
  var options =
   {
     "method"  : "post",
     "payload" : "message=" + message,
     "headers" : {"Authorization" : "Bearer "+ token}
     
   };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify",options);
}

/*#フォームのURLを取得する
function getFormUrl() {
  　//入力するメッセージボックス
  var formUrl = Browser.inputBox("こんにちは、あさののアンケートリマインド自動化プログラムです", "アンケートのリンクをコピペして下さい", Browser.Buttons.OK_CANCEL);
  if (formUrl != "cancel"){
    Browser.msgBox('入力されたのは以下のURLです、LINEで送信します' + formUrl);
  return formUrl;
  }else{
    return 0;
  }
}
*/


/*
//#メアドを名前に変換(ワンチャン漢字できたらここ消えるかも)
function getYetNames(){
  //#メアドリスト呼び出し
  var yetMails = getYetMails();
  Logger.log(yetMails);
  var yetNames = [];
  for(var i = 0;i < yetMails.length;i++){
    yetNames[i] = yetMails[i].replace('@aiesec.jp','');
  }
  Logger.log(yetNames);
  return yetNames;
}
*/

////以下より通知プログラム
  //メール本文リストを取得する※LINEbotでの通知のがよき？
  var docTest = DocumentApp.openById("1GqyGflwErlWnTErOqSfRCundksoYX6lUXD6WKgielRg");
  //Logger.log(docTest);//jsでいうconsole.log
  var strDoc=docTest.getBody().getText();
  
  
  //全未回答者に煽りメール送るめう
  /*
  for(var i = 0;i < yetMails.length; i++) {
  
    
    var strSubject = "アンケート未回答です";//メール文面のタイトル
    //メール文面の内容※ここ（””の中）にアンケートのリンクぶち込んでください
    var strBody = strDoc;
    var strFrom = "yuya.asano@aiesec.jp";//送信元メールアドレス
    var strSender = "ße-net";//送信者名  
    
      GmailApp.sendEmail(//gmailからメール送信
        yetMails[i],
        strSubject,
        strBody,
        {
          from:strFrom,
          name:strSender
        }      
    );
  }
  */