var SHEET_URL = "";
var SHEET_NAME = "BOT登録";

function SendMSG2() {
  var start = new Date(); //今日の日付
  var client = ChatWorkClient.factory({token: ""}); // Chatwork API
  var ss = SpreadsheetApp.openByUrl(SHEET_URL); //スプレッドシート取得
  var sheet = ss.getSheetByName(SHEET_NAME);　//シート取得
  var X = new Date();
  
  var calendar = CalendarApp.getCalendarsByName('日本の祝日')[0]; //祝日カレンダーを取得
  var event = calendar.getEventsForDay(start)[0];　//祝日カレンダーの今日のイベントを取得
  var lastRow = sheet.getLastRow() //最終行を取得
  var today = start.getDay() //今日の曜日を数字で取得（日曜＝0として0～6）
  var NowHour = start.getHours() //現在の時間（●時）を取得
  var ClmnG = sheet.getRange(2, 7, sheet.getLastRow() - 1).getValues();　//G列のデータ（送信曜日）を二元配列で取得
  var ClmnE = sheet.getRange(2, 5, sheet.getLastRow() - 1).getValues();　//E列のデータ（送信時間）を二元配列で取得
  var ClmnF = sheet.getRange(2, 6, sheet.getLastRow() - 1).getValues();　//F列のデータ（祝日の場合どうするか）を二元配列で取得
  var ClmnB = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues();　//B列のデータ（RoomID）を二元配列で取得
  var ClmnC = sheet.getRange(2, 3, sheet.getLastRow() - 1).getValues();　//C列（送付文言）のデータを二元配列で取得
  
　Logger.log(start);
  
  if (event == undefined //今日は営業日（祝日カレンダーの今日のイベントがない）
      ) 
   {for(var i=0;i<lastRow;i++)
   {if ( 
  ClmnG[i] == today //今日は通知曜日
  && ClmnE[i] == NowHour //今は通知時間
  ) 
{
    var Body = '' //謎だけどbody内に''がないとだめらしい
    Body += ClmnC[i]
    
    client.sendMessage({
    room_id:ClmnB[i], // room ID
    body:Body}); // message
  }
   }}
  
  if (event == undefined //今日は営業日（祝日カレンダーの今日のイベントがない）
      ) 
    { var Xn = []; //Xnを配列として定義
     Xn[0] = 0
     
     var Xday = []; //Xdayを配列として定義
     
  { for(var j=1;j<10;j++)
  { X.setDate(X.getDate() + 1); //翌日以降7日間の日付
     Xday[j] = X.getDay() //翌日以降7日間の曜日を数字で取得（日曜＝0として0～6）
     var eventX = calendar.getEventsForDay(X)[0];　//祝日カレンダーのXのイベントを取得
   if(eventX == 'CalendarEvent' //指定日は祝日（祝日カレンダーの指定日のイベントがある）
      || Xday[j] == 0 //指定日は日曜
      || Xday[j] == 6 //指定日は土曜
      )
   { Xn[j] = Xn[j-1]+1
   }
   else{
    break
   } } 
 　var XnN = Xn.length
  Logger.log(Xn);
  Logger.log(Xday);
   if(XnN > 1){for(var k=1;k<XnN;k++)
     {for(var l=0;l<lastRow;l++)
   {if (
  ClmnF[l] == '前営業日に投稿'
  && ClmnG[l] == Xday[k] //指定日は通知曜日
  && ClmnE[l] == NowHour //今は通知時間
   ) 
{   
    var BodyX = "※次回投稿日が祝日なので前倒しました※\n"
    BodyX += ClmnC[l]
    
    client.sendMessage({
    room_id:ClmnB[l], // room ID
    body:BodyX}); // message
}}}}}}

  
  var end = new Date();
  var span_sec = (end - start)/1000;
  Logger.log("処理時間は " + span_sec + " 秒でした" );   
  
}
