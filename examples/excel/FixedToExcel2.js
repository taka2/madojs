// �Œ蒷�t�@�C����1���R�[�h�̒������擾����
var recordLength = -1;
File.open(ARGV[0], "r", function(file) {
  recordLength = file.readLine().length;
});

var fieldInfo = [];
for(var i=0; i<recordLength; i++) {
  fieldInfo.push([i, Excel.COLUMN_TYPE_TEXT]);
}

// �Œ蒷�t�@�C������Excel�t�@�C�����쐬����B
Excel.convertFromFixed(ARGV[0], fieldInfo);

print("done.");
