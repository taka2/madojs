// Excel�t�@�C������V�[�g���𒊏o���A�e�L�X�g�t�@�C���ɏ����o���B
var inFileName = ARGV[0];
var outFileName = inFileName + ".txt";

Excel.openReadonly(inFileName, function(excel) {
  File.open(outFileName, "w", function(outfile) {
    excel.each(function(sheet) {
      outfile.puts(sheet.getName());
    });
  });
});

print("done.");