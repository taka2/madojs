Excel.create(function(excel) {
  var sheet = excel.getSheetByName("Sheet1");
  sheet.setValue(1, 1, 'hoge');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 50);

  sheet.setArrayValue(3, ["h1","h2","h3","---"]);
  sheet.setArrayValue(4, [1,2,3,"abc"]);
  sheet.setArrayValue(5, [4,5,6,"def"]);

  sheet.setAutoFilter(3, 1);
  sheet.setPageOrientation(Excel.PAGE_ORIENTATION_LANDSCAPE);
  sheet.setPageFitToPages(-1, -1);
  sheet.setPageZoom(80);
  excel.saveAs(SpecialFolders.getDesktop() + "\\test.xls");
});
