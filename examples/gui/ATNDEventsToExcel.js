function button1_Click() {
  HTTP.start("api.atnd.org", 80, function(http) {
    // ATND API�Ŗ{���̃C�x���g������
    var currentDateYYYYMMDD = document.form1.target_date.value;
    var atndResponse = JSON.parse(http.get("/events/?ymd=" + currentDateYYYYMMDD + "&count=100&format=json"));
    var resultCount = atndResponse.results_returned;
  
    // �o�̓t�@�C����
    var outputName = "ATND��������" + currentDateYYYYMMDD;
    //var outputFileName = File.realpath(outputName + ".xls", Dir.getParentDir(__FILE__));
    var outputFileName = outputName + ".xls";
  
    Excel.create(function(excel) {
      // �V�[�g�̎擾
      var sheet1 = excel.addSheet(outputName);
  
      // �w�b�_�̏����o��
      sheet1.setArrayValue(1, ["�C�x���gID", "�^�C�g��", "�J�n����", "�I������", "���"]);
  
      // �f�[�^�̏����o��
      var currentRow = 2;
      for(var i=0; i<resultCount; i++) {
        // ���X�|���X�f�[�^�̐��`�Ȃ�
        var eventObj = atndResponse.events[i];
        var no = i+1;
        var eventUrl = eventObj.event_url;
        var title = eventObj.title;
        var endedAt = eventObj.ended_at;
        if(endedAt) {
          endedAt = endedAt.replace("T", " ");
          endedAt = endedAt.replace("+09:00", "");
        }
        var eventId = eventObj.event_id;
        var startedAt = eventObj.started_at;
        if(startedAt) {
          startedAt = startedAt.replace("T", " ");
          startedAt = startedAt.replace("+09:00", "");
        }
        var place = eventObj.place;
  
        // �Z���ւ̏�������
        sheet1.addHyperLink(currentRow, 1, eventUrl, eventId+"");
        sheet1.setArrayValue(currentRow, [title, startedAt, endedAt, place], 2);
  
        currentRow++;
      }
  
      // Excel�V�[�g�̊e�풲��
      excel.removeSheets(3);
      for(var i=1; i<=5; i++) {
        sheet1.setBackgroundColor(1, i, Excel.COLOR_INDEX_37);
      }
      sheet1.autoFit();
      sheet1.drawBorder();
  
      // Excel�t�@�C���֕ۑ�
      excel.saveAs(outputFileName);
    });
  });
}