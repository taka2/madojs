Excel.create(function(excel) {
  var sheet = excel.getSheetByName("Sheet1");
  sheet.setValue(1, 1, 'hoge');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 50);
  excel.saveAs(SpecialFolders.getDesktop() + "\\test.xlsx");
});
