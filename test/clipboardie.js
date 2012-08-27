TestCase("ClipboardIE Test", {
  testHankaku: function() {
    var clip = new ClipboardIE();
    clip.set("abcde");
    assertEquals("abcde", clip.get());
    clip.close();
  },
  testZenkaku: function() {
    var clip = new ClipboardIE();
    clip.set("‚ ‚¢‚¤‚¦‚¨");
    assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
    clip.close();
  }
});
