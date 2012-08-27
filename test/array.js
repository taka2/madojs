TestCase("Array Test", {
  testForEach: function() {
    assertEquals([2, 4, 6].toString(), [1, 2, 3].forEach(function(x) {return x * 2;}).toString());
  },
  testEach: function() {
    assertEquals([2, 4, 6].toString(), [1, 2, 3].each(function(x) {return x * 2;}).toString());
  },
  testInclude: function() {
    assertEquals(true, [2, 4, 6].include(2));
  },
  testInclude2: function() {
    assertEquals(false, ["a", "b", "c"].include("d"));
  }
});
