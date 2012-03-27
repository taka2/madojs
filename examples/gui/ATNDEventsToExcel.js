function button1_Click() {
  HTTP.start("api.atnd.org", 80, function(http) {
    // ATND APIで本日のイベントを検索
    var currentDateYYYYMMDD = document.form1.target_date.value;
    var atndResponse = JSON.parse(http.get("/events/?ymd=" + currentDateYYYYMMDD + "&count=100&format=json"));
    var resultCount = atndResponse.results_returned;
  
    // 出力ファイル名
    var outputName = "ATND検索結果" + currentDateYYYYMMDD;
    //var outputFileName = File.realpath(outputName + ".xls", Dir.getParentDir(__FILE__));
    var outputFileName = outputName + ".xls";
  
    Excel.create(function(excel) {
      // シートの取得
      var sheet1 = excel.addSheet(outputName);
  
      // ヘッダの書き出し
      sheet1.setArrayValue(1, ["イベントID", "タイトル", "開始日時", "終了日時", "会場"]);
  
      // データの書き出し
      var currentRow = 2;
      for(var i=0; i<resultCount; i++) {
        // レスポンスデータの整形など
        var eventObj = atndResponse.events[i];
        var no = i+1;
        var eventUrl = eventObj.event_url;
        var title = eventObj.title;
        var endedAt = eventObj.ended_at;
        if(endedAt) {
          endedAt = endedAt.replace("T", " ");
          endedAt = endedAt.replace("+09:00", "");
        }
        var eventId = eventObj.event_id;
        var startedAt = eventObj.started_at;
        if(startedAt) {
          startedAt = startedAt.replace("T", " ");
          startedAt = startedAt.replace("+09:00", "");
        }
        var place = eventObj.place;
  
        // セルへの書き込み
        sheet1.addHyperLink(currentRow, 1, eventUrl, eventId+"");
        sheet1.setArrayValue(currentRow, [title, startedAt, endedAt, place], 2);
  
        currentRow++;
      }
  
      // Excelシートの各種調整
      excel.removeSheets(3);
      for(var i=1; i<=5; i++) {
        sheet1.setBackgroundColor(1, i, Excel.COLOR_INDEX_37);
      }
      sheet1.autoFit();
      sheet1.drawBorder();
  
      // Excelファイルへ保存
      excel.saveAs(outputFileName);
    });
  });
}