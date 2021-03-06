TestCase("ClipboardIE Test", {
  testOpen: function() {
    ClipboardIE.open(function(clip) {
      clip.set("12345");
      assertEquals("12345", clip.get());
    });
  },
  testHankaku: function() {
    var clip = new ClipboardIE();
    clip.set("abcde");
    assertEquals("abcde", clip.get());
    clip.close();
  },
  testZenkaku: function() {
    var clip = new ClipboardIE();
    clip.set("あいうえお");
    assertEquals("あいうえお", clip.get());
    clip.close();
  }
});
