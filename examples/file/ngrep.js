// inFileNameファイルの中で、patternにマッチする行と、その前後numLinesを出力する。

if(ARGV.length < 3) {
  print("Usage: cscript grep.wsf [pattern] [numLines] [inFileName]");
  exit(1);
}
var pattern = ARGV[0];
var numLines = ARGV[1];
var inFileName = ARGV[2];

var buffer = [];
var outputLines = 0;
File.open(inFileName, "r", function(infile) {
  var lineNumber = 1;
  infile.each(function(line) {
    // マッチ行、および、マッチ行の前numLinesの出力
    if(line.match(pattern)) {
      var bufferLength = buffer.length;
      for(var i=0; i<bufferLength; i++) {
        printLine("-", (lineNumber - bufferLength + i), buffer[i]);
      }
      printLine("*", lineNumber, line);
      outputLines = numLines;
    } else if(outputLines > 0) {
      // マッチ行の後ろnumLinesの出力
      printLine("+", lineNumber, line);
      outputLines--;
    }

    // バッファの操作
    buffer.push(line);
    if(buffer.length > numLines) {
      buffer.shift();
    }

    lineNumber++;
  });
});

function printLine(outputType, lineNumber, content) {
  var strLineNumber = lineNumber + "";
  print(outputType + strLineNumber.lpad(" ", 5) + ":" + content);
}

