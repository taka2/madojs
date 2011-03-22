Test.String = {};

Test.String.testTrim = function() {
  testUtil.assertEquals("@abc@", "@" + "   	abc			".trim() + "@");
}
