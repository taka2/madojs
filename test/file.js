Test.File = {};

Test.File.testFile = function() {
  // Create File
  File.open("test.txt", "w", function(file) {
    file.puts("hoge");
  });
  testUtil.assertEquals(true, File.exist("test.txt"));

  // Rename File
  File.rename("test.txt", "test2.txt");
  testUtil.assertEquals(false, File.exist("test.txt"));
  testUtil.assertEquals(true, File.exist("test2.txt"));

  // Copy File
  File.copy("test2.txt", "test3.txt");
  testUtil.assertEquals(true, File.exist("test2.txt"));
  testUtil.assertEquals(true, File.exist("test3.txt"));

  // File Size
  testUtil.assertEquals(6, File.size("test2.txt"));
  testUtil.assertEquals(6, File.size("test3.txt"));

  // Open File
  File.open("test2.txt", "r", function(file) {
    file.each(function(line) {
      testUtil.assertEquals("hoge", line);
    });
  });

  // Ext Name
  testUtil.assertEquals(".txt", File.extname("test2.txt"));

  // Delete Files
  File.unlink("test2.txt");
  File.unlink("test3.txt");
  testUtil.assertEquals(false, File.exist("test2.txt"));
  testUtil.assertEquals(false, File.exist("test3.txt"));
};
