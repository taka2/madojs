Test.Dir = {};

Test.Dir.testDir = function() {
  // Create Directory
  Dir.mkdir("hoge");
  testUtil.assertEquals(true, Dir.exist("hoge"));

  // Delete Directory
  File.unlink("hoge");
  testUtil.assertEquals(false, Dir.exist("hoge"));
};
