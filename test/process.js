Test.Process = {};

Test.Process.testExec = function() {
  Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_MIN, true);
  var filePath = File.realpath("test.txt", Dir.getwd());
  testUtil.assertEquals(true, File.exist(filePath));
  File.unlink(filePath);
};

Test.Process.testExec2 = function() {
  Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_NORMAL, true);
  var filePath = File.realpath("test.txt", Dir.getwd());
  testUtil.assertEquals(true, File.exist(filePath));
  File.unlink(filePath);
};

Test.Process.testExec3 = function() {
  Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_MAX, true);
  var filePath = File.realpath("test.txt", Dir.getwd());
  testUtil.assertEquals(true, File.exist(filePath));
  File.unlink(filePath);
};
