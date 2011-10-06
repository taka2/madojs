if(Excel.available()) {
  Test.Excel = {};

  Test.Excel.testExcel = function() {
    Excel.create(function(excel) {
      var sheet = excel.getSheetByIndex(1);
      sheet.setValue(1, 1, "abc");
      excel.saveAs("test.xls");
    });
    testUtil.assertEquals(true, File.exist("test.xls"));

    File.unlink("test.xls");
  };
}