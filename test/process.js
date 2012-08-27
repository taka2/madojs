TestCase("Process Test", {
  testExec: function() {
    Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_MIN, true);
    var filePath = File.realpath("test.txt", Dir.getwd());
    assertEquals(true, File.exist(filePath));
    File.unlink(filePath);
  },
  testExec2: function() {
    Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_NORMAL, true);
    var filePath = File.realpath("test.txt", Dir.getwd());
    assertEquals(true, File.exist(filePath));
    File.unlink(filePath);
  },
  testExec3: function() {
    Process.exec("cmd", ["/c", "echo abc > test.txt"], Const.WINDOW_STYLE_MAX, true);
    var filePath = File.realpath("test.txt", Dir.getwd());
    assertEquals(true, File.exist(filePath));
    File.unlink(filePath);
  }
});
