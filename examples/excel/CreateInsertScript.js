// Excel�t�@�C������INSERT���̏����ꂽSQL�t�@�C�����쐬����B
var inFileName = ARGV[0];

Excel.openReadonly(inFileName, function(excel) {
  // �S�V�[�g�����Ԃɏ�������
  excel.each(function(sheet) {
    var outFileName = inFileName + "_" + sheet.getName() + ".sql";
    processSheet(sheet, outFileName);
  });
});

print("done.");

function processSheet(sheet, outFileName) {
  // ���ږ��̓ǂݎ��
  var columns = [];
  var i = 1;
  while(true) {
    var columnName = sheet.getValue(1, i);
    if(columnName) {
      columns.push(columnName);
    } else {
      break;
    }
    i++;
  }
  var columnCount = columns.length;
  var commaSeparatedColumnNames = columns.join(",");

  // �f�[�^�̓ǂݎ���SQL�t�@�C���ւ̏����o��
  File.open(outFileName, "w", function(outfile) {
    var currentRow = 2;
    while(true) {
      var dataList = [];

      // �擪�J�����̃f�[�^����̏ꍇ���[�v���I������
      var isDataExists = sheet.getValue(currentRow, 1);
      if(!isDataExists) {
        break;
      }

      // �f�[�^��ǂݍ���Ńt�@�C���ɏ����o��
      for(i=1; i<=columnCount; i++) {
        var data = sheet.getValue(currentRow, i);
        if(data) {
          dataList.push(data);
        } else {
          dataList.push("NULL");
        }
      }

      outfile.puts("INSERT INTO " + sheet.getName() + "(" + commaSeparatedColumnNames + ") VALUES (" + dataList.join(",") + ");");
      currentRow++;
    }
  });
}
