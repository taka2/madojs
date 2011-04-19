Test.String = {};

Test.String.testTrim = function() {
  testUtil.assertEquals("@abc@", "@" + "   	abc			".trim() + "@");
};

Test.String.testStartsWith = function() {
  testUtil.assertEquals(true, "abc".startsWith("a"));
};

Test.String.testStartsWith2 = function() {
  testUtil.assertEquals(false, "abc".startsWith("b"));
};

Test.String.testEndsWith = function() {
  testUtil.assertEquals(true, "abc".endsWith("c"));
};

Test.String.testEndsWith = function() {
  testUtil.assertEquals(false, "abc".endsWith("b"));
};

Test.String.testIsZenkaku = function() {
  testUtil.assertEquals(true, "‚ ".isZenkaku());
};

Test.String.testIsZenkaku2 = function() {
  testUtil.assertEquals(false, "±".isZenkaku());
};

Test.String.testIsHankakuKana = function() {
  testUtil.assertEquals(true, "§".isHankakuKana());
};

Test.String.testIsHankakuKana = function() {
  testUtil.assertEquals(false, "ƒA".isHankakuKana());
};

Test.String.testSubstringb = function() {
  testUtil.assertEquals("±", "±ƒA1‚P".substringb(0, 2));
};

Test.String.testSubstringb2 = function() {
  testUtil.assertEquals("±ƒA", "±ƒA1‚P".substringb(0, 3));
};

Test.String.testSubstringb3 = function() {
  testUtil.assertEquals("±ƒA1", "±ƒA1‚P".substringb(0, 4));
};
