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
