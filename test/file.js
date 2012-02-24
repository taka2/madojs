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

  File.open("test2.txt", "r", function(file) {
    testUtil.assertEquals("hoge", file.readLine());
  });

  // Ext Name
  testUtil.assertEquals(".txt", File.extname("test2.txt"));

  // Delete Files
  File.unlink("test2.txt");
  File.unlink("test3.txt");
  testUtil.assertEquals(false, File.exist("test2.txt"));
  testUtil.assertEquals(false, File.exist("test3.txt"));
};

Test.File.testDirname = function() {
  // Test Case from http://pubs.opengroup.org/onlinepubs/009695399/functions/dirname.html
  testUtil.assertEquals("/usr", File.dirname("/usr/lib"));
  testUtil.assertEquals("/", File.dirname("/usr/"));
  testUtil.assertEquals(".", File.dirname("usr"));
  testUtil.assertEquals("/", File.dirname("/"));
  testUtil.assertEquals(".", File.dirname("."));
  testUtil.assertEquals(".", File.dirname(".."));

  // Test Case from http://doc.ruby-lang.org/ja/1.9.2/class/File.html
  testUtil.assertEquals("dir", File.dirname("dir/file.ext"));
  testUtil.assertEquals(".", File.dirname("file.ext"));
  testUtil.assertEquals("foo", File.dirname("foo/bar/"));
  testUtil.assertEquals("foo", File.dirname("foo//bar"));

  // Original Test Case
  testUtil.assertEquals("/", File.dirname("////aaa"));
}

Test.File.testGetShortName = function() {
  testUtil.assertEquals("MADO-D~1.JS", File.getShortName("mado-debug.js"));
  File.open("mado-debug.js", "r", function(file) {
    testUtil.assertEquals("MADO-D~1.JS", file.getShortName());
  });
}