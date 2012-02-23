// inFileNameファイルの中で、patternにマッチする行と、その前後numLinesを出力する。

if(ARGV.length < 2) {
  print("Usage: cscript grep.wsf [pattern] [inFileName]");
  exit(1);
}
var pattern = ARGV[0];
var inFileName = ARGV[1];

File.open(inFileName, "r", function(infile) {
  var lineNumber = 1;
  infile.each(function(line) {
    if(line.match(pattern)) {
      print(lineNumber + ":" + line);
    }
    lineNumber++;
  });
});

