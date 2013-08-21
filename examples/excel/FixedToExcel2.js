// 固定長ファイルの1レコードの長さを取得する
var recordLength = -1;
File.open(ARGV[0], "r", function(file) {
  recordLength = file.readLine().length;
});

var fieldInfo = [];
for(var i=0; i<recordLength; i++) {
  fieldInfo.push([i, Excel.COLUMN_TYPE_TEXT]);
}

// 固定長ファイルからExcelファイルを作成する。
Excel.convertFromFixed(ARGV[0], fieldInfo);

print("done.");
