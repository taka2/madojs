TestCase("ClipboardExcel Test", {
  testHankaku: function() {
    var clip = new ClipboardExcel();
    clip.set("abcde");
    assertEquals("abcde", clip.get());
    clip.close();
  },
  testZenkaku: function() {
    var clip = new ClipboardExcel();
    clip.set("‚ ‚¢‚¤‚¦‚¨");
    assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
    clip.close();
  }
});
