// 1�s1���`���̃e�L�X�g�t�@�C������A���̃e�L�X�g���V�[�g���Ɏ���Excel�t�@�C�����쐬����B

Excel.create(function(excel) {
  File.open(ARGV[0], "r", function(infile) {
    infile.each(function(line) {
      if(line.trim().length != 0) {
        // ��s����Ȃ�������V�[�g�ǉ�
        excel.addSheet(line);
      }
    });
  });
  // �f�t�H���g�̃V�[�g���폜
  excel.removeSheets(3);
  // Excel�t�@�C���Ƃ��ĕۑ�
  excel.saveAs(ARGV[0] + ".xls");
});

print("done.");