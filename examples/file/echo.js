// �t�@�C���̓��e��\�����܂��B
var inFileName = ARGV[0];
File.open(inFileName, "r", function(infile) {
  var str = "";
  while(!infile.eof()) {
    str += infile.readLine();
  }

  print(str);
});

