// examples�ȉ��̃t�@�C�������f�X�N�g�b�v\list.txt�ɏo�͂���B

File.open(SpecialFolders.getDesktop() + "\\list.txt", "w", function(file) {
  Dir.find("..", ".*", function(fileName) {
    file.puts(fileName);
  });
});
