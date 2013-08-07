TestCase("ClipboardExcel Test", {
  testOpen: function() {
    if(Excel.available()) {
      ClipboardExcel.open(function(clip) {
        clip.set("12345");
        assertEquals("12345", clip.get());
      });
    }
  },
  testHankaku: function() {
    if(Excel.available()) {
      var clip = new ClipboardExcel();
      clip.set("abcde");
      assertEquals("abcde", clip.get());
      clip.close();
    }
  },
  testZenkaku: function() {
    if(Excel.available()) {
      var clip = new ClipboardExcel();
      clip.set("‚ ‚¢‚¤‚¦‚¨");
      assertEquals("‚ ‚¢‚¤‚¦‚¨", clip.get());
      clip.close();
    }
  }
});
