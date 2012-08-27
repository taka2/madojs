TestCase("ClipboardExcel Test", {
  testOpen: function() {
    ClipboardExcel.open(function(clip) {
      clip.set("12345");
      assertEquals("12345", clip.get());
    });
  },
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
