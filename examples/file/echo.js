// ファイルの内容を表示します。
var inFileName = ARGV[0];
File.open(inFileName, "r", function(infile) {
  var str = "";
  while(!infile.eof()) {
    str += infile.readLine();
  }

  print(str);
});

