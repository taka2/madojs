// 01�`30�܂ł̃t�H���_���쐬����B

var NUM_FOLDERS = 30;

for(var i=1; i<=NUM_FOLDERS; i++) {
  var num = (i+"").lpad("0", 2);
  Dir.mkdir(num)
}