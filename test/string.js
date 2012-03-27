TestCase("String Test", {
  testTrim: function() {
    assertEquals("@abc@", "@" + "   	abc			".trim() + "@");
  },
  testStartsWith: function() {
    assertEquals(true, "abc".startsWith("a"));
  },
  testStartsWith2: function() {
    assertEquals(true, "abc".startsWith("ab"));
  },
  testStartsWith3: function() {
    assertEquals(false, "abc".startsWith("b"));
  },
  testStartsWith4: function() {
    assertEquals(false, "abc".startsWith("bc"));
  },
  testEndsWith: function() {
    assertEquals(true, "abc".endsWith("c"));
  },
  testEndsWith2: function() {
    assertEquals(true, "abc".endsWith("bc"));
  },
  testEndsWith3: function() {
    assertEquals(false, "abc".endsWith("b"));
  },
  testEndsWith4: function() {
    assertEquals(false, "abc".endsWith("ab"));
  },
  testIsZenkaku: function() {
    assertEquals(true, "‚ ".isZenkaku());
  },
  testIsZenkaku2: function() {
    assertEquals(false, "±".isZenkaku());
  },
  testIsHankakuKana: function() {
    assertEquals( true, "±".isHankakuKana());
  },
  testIsHankakuKana2: function() {
    assertEquals(false, "ƒA".isHankakuKana());
  },
  testSubstringb: function() {
    assertEquals("±", "±ƒA1‚P".substringb(0, 2));
  },
  testSubstringb2: function() {
    assertEquals("±ƒA", "±ƒA1‚P".substringb(0, 3));
  },
  testSubstringb3: function() {
    assertEquals("±ƒA1", "±ƒA1‚P".substringb(0, 4));
  },
  testLpad: function() {
    assertEquals("00012", "12".lpad("0", 5));
  }
});
