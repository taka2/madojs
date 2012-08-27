TestCase("Dir Test", {
  testDir: function() {
    // Create Directory
    Dir.mkdir("hoge");
    assertEquals(true, Dir.exist("hoge"));

    // Delete Directory
    File.unlink("hoge");
    assertEquals(false, Dir.exist("hoge"));
  }
});
