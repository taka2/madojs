Test.Clipboard = {};

Test.Clipboard.testHankaku = function() {
  var clip = new Clipboard();
  clip.set("abcde");
  testUtil.assertEquals("abcde", clip.get());
  clip.close();
};

Test.Clipboard.testHankaku2 = function() {
  Clipboard.open(function(clip) {
    clip.set("fghij");
    testUtil.assertEquals("fghij", clip.get());
  });
};

Test.Clipboard.testZenkaku = function() {
  var clip = new Clipboard();
  clip.set("‚ ‚¢‚¤‚¦‚¨");
  testUtil.assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
  clip.close();
};

Test.Clipboard.testZenkaku2 = function() {
  Clipboard.open(function(clip) {
    clip.set("‚©‚«‚­‚¯‚±");
    testUtil.assertEquals("‚©‚«‚­‚¯‚±", clip.get());
  });
};
