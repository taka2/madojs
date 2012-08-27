// examples以下のファイル名をデスクトップ\list.txtに出力する。

File.open(SpecialFolders.getDesktop() + "\\list.txt", "w", function(file) {
  Dir.find("..", ".*", function(fileName) {
    file.puts(fileName);
  });
});
