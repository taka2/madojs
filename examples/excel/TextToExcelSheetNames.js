// 1行1件形式のテキストファイルから、そのテキストをシート名に持つExcelファイルを作成する。

Excel.create(function(excel) {
  File.open(ARGV[0], "r", function(infile) {
    infile.each(function(line) {
      if(line.trim().length != 0) {
        // 空行じゃなかったらシート追加
        excel.addSheet(line);
      }
    });
  });
  // デフォルトのシートを削除
  excel.removeSheets(3);
  // Excelファイルとして保存
  excel.saveAs(ARGV[0] + ".xls");
});

print("done.");