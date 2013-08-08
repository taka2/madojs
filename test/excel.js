TestCase("Excel Test", {
  testExcel: function() {
    if(Excel.available()) {
      Excel.create(function(excel) {
        var sheet = excel.getSheetByIndex(1);
        sheet.setValue(1, 1, "abc");
        assertEquals("abc", sheet.getValue(1, 1));
        sheet.setArrayValue(2, [1.23, true, "def"], 3);
        assertEquals(1.23, sheet.getValue(2, 3));
        assertEquals(true, sheet.getValue(2, 4));
        assertEquals("def", sheet.getValue(2, 5));
        excel.saveAs("test.xls");
      });
      assertEquals(true, File.exist("test.xls"));

      File.unlink("test.xls");
    }
  },
  testArray2dToSafeArray2d: function() {
    if(Excel.available()) {
      var actual = new VBArray(Excel.array2dToSafeArray2d([[1,2,3],[4,5,6]]));
      var expected = [[1,2,3],[4,5,6]];
      for(var i=1; i<=(actual.ubound(1)); i++) {
        for(var j=1; j<=(actual.ubound(2)); j++) {
          assertEquals(actual.getItem(i, j), expected[i-1][j-1]);
        }
      }
    }
  }
});
