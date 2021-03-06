TestCase("Clipboard Test", {
  testHankaku: function() {
    var clip = new Clipboard();
    clip.set("abcde");
    assertEquals("abcde", clip.get());
    clip.close();
  },
  testHankaku2: function() {
    Clipboard.open(function(clip) {
      clip.set("fghij");
      assertEquals("fghij", clip.get());
    });
  },
  testZenkaku: function() {
    var clip = new Clipboard();
    clip.set("あいうえお");
    assertEquals("あいうえお", clip.get());
    clip.close();
  },
  testZenkaku2: function() {
    Clipboard.open(function(clip) {
      clip.set("かきくけこ");
      assertEquals("かきくけこ", clip.get());
    });
  }
});
