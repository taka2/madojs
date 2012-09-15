TestCase("File Test", {
  testFile: function() {
    // Create File
    File.open("test.txt", "w", function(file) {
      file.puts("hoge");
    });
    assertEquals(true, File.exist("test.txt"));

    // Rename File
    File.rename("test.txt", "test2.txt");
    assertEquals(false, File.exist("test.txt"));
    assertEquals(true, File.exist("test2.txt"));

    // Copy File
    File.copy("test2.txt", "test3.txt");
    assertEquals(true, File.exist("test2.txt"));
    assertEquals(true, File.exist("test3.txt"));

    // File Size
    assertEquals(6, File.size("test2.txt"));
    assertEquals(6, File.size("test3.txt"));

    // Open File
    File.open("test2.txt", "r", function(file) {
      file.each(function(line) {
        assertEquals("hoge", line);
      });
    });

    // Open File
    File.open("test2.txt", "r", function(file) {
      assertEquals("hoge", file.readLine());
      assertEquals(true, file.eof());
    });

    File.open("test2.txt", "r", function(file) {
      assertEquals("hoge", file.readLine());
    });

    File.open("test2.txt", "r", function(file) {
      assertEquals("hoge\r\n", file.read());
    });

    // Ext Name
    assertEquals(".txt", File.extname("test2.txt"));

    // Delete Files
    File.unlink("test2.txt");
    File.unlink("test3.txt");
    assertEquals(false, File.exist("test2.txt"));
    assertEquals(false, File.exist("test3.txt"));
  },
  testDirname: function() {
    // Test Case from http://pubs.opengroup.org/onlinepubs/009695399/functions/dirname.html
    assertEquals("/usr", File.dirname("/usr/lib"));
    assertEquals("/", File.dirname("/usr/"));
    assertEquals(".", File.dirname("usr"));
    assertEquals("/", File.dirname("/"));
    assertEquals(".", File.dirname("."));
    assertEquals(".", File.dirname(".."));

    // Test Case from http://doc.ruby-lang.org/ja/1.9.2/class/File.html
    assertEquals("dir", File.dirname("dir/file.ext"));
    assertEquals(".", File.dirname("file.ext"));
    assertEquals("foo", File.dirname("foo/bar/"));
    assertEquals("foo", File.dirname("foo//bar"));

    // Original Test Case
    assertEquals("/", File.dirname("////aaa"));
  },
  testGetShortName: function() {
    assertEquals("MADO-D~1.JS", File.getShortName("mado-debug.js"));
    File.open("mado-debug.js", "r", function(file) {
      assertEquals("MADO-D~1.JS", file.getShortName());
    });
  }
});
