// 01～30までのフォルダを作成する。

var NUM_FOLDERS = 30;

for(var i=1; i<=NUM_FOLDERS; i++) {
  var num = (i+"").lpad("0", 2);
  Dir.mkdir(num)
}