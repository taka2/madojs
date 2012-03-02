if(Excel.available()) {
  Test.Excel = {};

  Test.Excel.testExcel = function() {
    Excel.create(function(excel) {
      var sheet = excel.getSheetByIndex(1);
      sheet.setValue(1, 1, "abc");
      testUtil.assertEquals("abc", sheet.getValue(1, 1));
      sheet.setArrayValue(2, [1.23, true, "def"], 3);
      testUtil.assertEquals(1.23, sheet.getValue(2, 3));
      testUtil.assertEquals(true, sheet.getValue(2, 4));
      testUtil.assertEquals("def", sheet.getValue(2, 5));
      excel.saveAs("test.xls");
    });
    testUtil.assertEquals(true, File.exist("test.xls"));

    File.unlink("test.xls");
  };
}