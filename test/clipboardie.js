Test.ClipboardIE = {};

Test.ClipboardIE.testHankaku = function() {
  var clip = new ClipboardIE();
  clip.set("abcde");
  testUtil.assertEquals("abcde", clip.get());
  clip.close();
};

Test.ClipboardIE.testZenkaku = function() {
  var clip = new ClipboardIE();
  clip.set("‚ ‚¢‚¤‚¦‚¨");
  testUtil.assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
  clip.close();
};
