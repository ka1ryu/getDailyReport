var toDay = new Date();
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("mail");
var sheetInst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instruction");
var lastRow = sheet.getLastRow();
var report = "";
//var roomID = sheetInst.getRange(2,1).getValue();
var Sub = sheetInst.getRange('B2').getValue();


function searchContactMail() {
  
  /* Gmailから特定条件のスレッドを検索しメールを取り出す */
  var strTerms = 'subject:'+Sub; //メール検索条件
  var numMailMax = 5; //取得するメール総数  
  var numMail = 1; //1度に取得するメール数
  var myThreads; //条件にマッチしたスレッドを取得、最大500通と決まっている
  var myMsgs; //スレッドからメールを取得する　→二次元配列で格納
  var valMsgs;
 
  var i = SpreadsheetApp.getActiveSheet().getLastRow();
 
  if(i<numMailMax) {  
    valMsgs = [];
    myThreads = GmailApp.search(strTerms, i, numMail); //条件にマッチしたスレッドを取得、最大500通と決まっている
    myMsgs = GmailApp.getMessagesForThreads(myThreads); //スレッドからメールを取得する　→二次元配列で格納
 
    /* 各メールから日時、送信元、件名、内容を取り出す*/
    for(var j = 0;j < myMsgs.length;j++){
      valMsgs[j] = [];
      valMsgs[j][0] = myMsgs[j][0].getDate();
      valMsgs[j][1] = myMsgs[j][0].getSubject();
      valMsgs[j][2] = myMsgs[j][0].getPlainBody();
      var report = valMsgs[0][2];
    }
 
    /* スプレッドシートに出力 */
    if(myMsgs.length>0){
      var insert = sheet.getRange(lastRow+1, 1, j, 3).setValues(valMsgs); //シートに貼り付け
      
      
      var client = ChatWorkClient.factory({token: 'd124ab756105618eadad7252cf6ef892'});　//チャットワークAPI
      client.sendMessage({
        room_id:31434119, //ルームID
        body:toDay + String.fromCharCode(10) + report});
    } 
  }
}

