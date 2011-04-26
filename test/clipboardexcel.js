if(Excel.available()) {
  Test.ClipboardExcel = {};

  Test.ClipboardExcel.testHankaku = function() {
    var clip = new ClipboardExcel();
    clip.set("abcde");
    testUtil.assertEquals("abcde", clip.get());
    clip.close();
  };

  Test.ClipboardExcel.testZenkaku = function() {
    var clip = new ClipboardExcel();
    clip.set("‚ ‚¢‚¤‚¦‚¨");
    testUtil.assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
    clip.close();
  };
}