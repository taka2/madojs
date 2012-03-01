// �{�����t��YYYYMMDD�`���Ŏ擾
var currentDateYYYYMMDD = new Date().getYYYYMMDD();

// ATND API�Ŗ{���̃C�x���g������
HTTP.start("api.atnd.org", 80, function(http) {
  var atndResponse = eval('(' + http.get("/events/?ymd=" + currentDateYYYYMMDD + "&count=100&format=json") + ')');
  var resultCount = atndResponse.results_returned;
  var outputFileName = File.realpath("TodaysEvents.xls", Dir.getParentDir(__FILE__));

  Excel.create(function(excel) {
    var sheet1 = excel.getSheetByIndex(0);

    // �w�b�_�̏����o��
    sheet1.setValue(1, 1, "�C�x���gID");
    sheet1.setValue(1, 2, "�^�C�g��");
    sheet1.setValue(1, 3, "�J�n����");
    sheet1.setValue(1, 4, "�I������");
    sheet1.setValue(1, 5, "���");

    // �f�[�^�̏����o��
    var currentRow = 2;
    for(var i=0; i<resultCount; i++) {
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

      sheet1.addHyperLink(currentRow, 1, eventUrl, eventId+"");
      sheet1.setValue(currentRow, 2, title);
      sheet1.setValue(currentRow, 3, startedAt);
      sheet1.setValue(currentRow, 4, endedAt);
      sheet1.setValue(currentRow, 5, place);

      currentRow++;
    }

    // �e�풲��
    for(var i=1; i<=5; i++) {
      sheet1.setBackgroundColor(1, i, Excel.COLOR_INDEX_37);
    }
    sheet1.autoFit();
    sheet1.drawBorder();

    // Excel�t�@�C���֕ۑ�
    excel.saveAs(outputFileName);
  });
});
