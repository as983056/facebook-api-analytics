//粉專資料
const PAGE_ID = 'YOUR_PAGE_ID';
const TOKEN = "YOUR_TOKEN";
const PAGE_ACCESS_TOKEN = 'YOUR_PAGE_ACCESS_TOKEN';
const FIELDSDATA = 'id,message,created_time,permalink_url,shares,reactions.limit(0).summary(total_count),comments.limit(0).summary(total_count)'
const FIELDSDATA2 = 'post_impressions_organic_unique'
  
function facebookdata() {
  var sheet_url = 'YOUR_SHEET_URL';   //試算表連結(可變更)
  var SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);        //找到試算表(固定)
  var sheet_name = '貼文數據';                                   //工作表名稱(工作表改名這邊要跟著動)
  var reserve_list = SpreadSheet.getSheetByName(sheet_name);    //找到工作表(固定)
  var reserve_list_row = reserve_list.getLastRow();             //工作表"列數"(固定)
  var hours = Utilities.formatDate(new Date(), "GMT+8", "HH")   //目前小時
  var mins = Utilities.formatDate(new Date(), "GMT+8", "mm")    //目前分鐘
  var test = '';

  //每天跨日執行
  if(hours == 0 & mins == 0){
    do{
      try{
        //建立Facebook 廣告類型API連結(固定)
        const facebookUrl = `https://graph.facebook.com/v13.0/${PAGE_ID}/posts?fields=${FIELDSDATA}&access_token=${TOKEN}&limit=100`;

        //獲得Facebook API抓取"資料"(固定)
        const encodedFacebookUrl = encodeURI(facebookUrl);
        const options = {
          'method' : 'get'
        };
        const fetchRequest = UrlFetchApp.fetch(encodedFacebookUrl, options);
        var results = JSON.parse(fetchRequest.getContentText());  //JSON資料轉換編碼
        results = results.data  //"擷取"API內"data"資料
        var rows = [],data;     //新增空陣列
        test = 'ok';            //完成
      }
      catch{
        test = '';              //失敗
      }
    }while(test == '');

    //"整理"JSON內data資料
    for (i = 0; i < results.length; i++) {
      data = results[i];

      var shares = 0;
      try{
        shares = data.shares.count;
      }
      catch{
        shares = 0;
      }
      var reactions = data.reactions.summary.total_count;
      var comments = data.comments.summary.total_count;

      rows.push([data.id, data.created_time, data.message, data.permalink_url, comments, reactions, shares]);
    }
    reserve_list.getRange(2, 2, reserve_list_row - 1, 7).setValue('');
    dataRange = reserve_list.getRange(2, 2, rows.length, 7);                       //放上整理過的資料到試算表(可變更)
    dataRange.setValues(rows);                                                     //放上整理過的資料到試算表(固定)

    /********************************************************************************************************************/
    reserve_list_row = reserve_list.getLastRow();
    var rows = [];

    //在每個貼文數據後面補個觸及
    for(a = 2; a <= reserve_list_row; a++){
      do{
        try{
          var page_post_id = reserve_list.getRange(a, 2).getValue();

          //建立Facebook 廣告類型API連結(固定)
          const facebookUrl = `https://graph.facebook.com/v13.0/${page_post_id}/insights?metric=${FIELDSDATA2}&access_token=${PAGE_ACCESS_TOKEN}`;

          //獲得Facebook API抓取"資料"(固定)
          const encodedFacebookUrl = encodeURI(facebookUrl);
          const options = {
            'method' : 'get'
          };
          const fetchRequest = UrlFetchApp.fetch(encodedFacebookUrl, options);
          var results = JSON.parse(fetchRequest.getContentText());  //JSON資料轉換編碼
          results = results.data  //"擷取"API內"data"資料
          data = results[0]
          test = 'ok';            //完成
        }
        catch{
          test = '';              //失敗
        }
      }while(test == '');

      //"整理"API內data資料
      rows.push([data.values[0].value])
    }
    reserve_list.getRange(2, 9, reserve_list_row - 1, 1).setValue('');
    dataRange2 = reserve_list.getRange(2, 9, rows.length, 1);                       //放上整理過的資料到試算表(可變更)
    dataRange2.setValues(rows);                                                     //放上整理過的資料到試算表(固定)
  }
}
