Test.Clipboard = {};

Test.Clipboard.testHankaku = function() {
  var clip = new Clipboard();
  clip.set("abcde");
  testUtil.assertEquals("abcde", clip.get());
};

Test.Clipboard.testZenkaku = function() {
  var clip = new Clipboard();
  clip.set("‚ ‚¢‚¤‚¦‚¨");
  testUtil.assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
};