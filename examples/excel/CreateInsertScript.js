// ExcelファイルからINSERT文の書かれたSQLファイルを作成する。
var inFileName = ARGV[0];

Excel.openReadonly(inFileName, function(excel) {
  // 全シートを順番に処理する
  excel.each(function(sheet) {
    var outFileName = inFileName + "_" + sheet.getName() + ".sql";
    processSheet(sheet, outFileName);
  });
});

print("done.");

function processSheet(sheet, outFileName) {
  // 項目名の読み取り
  var columns = [];
  var i = 1;
  while(true) {
    var columnName = sheet.getValue(1, i);
    if(columnName) {
      columns.push(columnName);
    } else {
      break;
    }
    i++;
  }
  var columnCount = columns.length;
  var commaSeparatedColumnNames = columns.join(",");

  // データの読み取りとSQLファイルへの書き出し
  File.open(outFileName, "w", function(outfile) {
    var currentRow = 2;
    while(true) {
      var dataList = [];

      // 先頭カラムのデータが空の場合ループを終了する
      var isDataExists = sheet.getValue(currentRow, 1);
      if(!isDataExists) {
        break;
      }

      // データを読み込んでファイルに書き出す
      for(i=1; i<=columnCount; i++) {
        var data = sheet.getValue(currentRow, i);
        if(data) {
          dataList.push(data);
        } else {
          dataList.push("NULL");
        }
      }

      outfile.puts("INSERT INTO " + sheet.getName() + "(" + commaSeparatedColumnNames + ") VALUES (" + dataList.join(",") + ");");
      currentRow++;
    }
  });
}
