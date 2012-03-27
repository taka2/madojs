// Excelファイルからシート名を抽出し、テキストファイルに書き出す。
var inFileName = ARGV[0];
var outFileName = inFileName + ".txt";

Excel.openReadonly(inFileName, function(excel) {
  File.open(outFileName, "w", function(outfile) {
    excel.each(function(sheet) {
      outfile.puts(sheet.getName());
    });
  });
});

print("done.");