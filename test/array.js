Test.Array = {};

Test.Array.testEach = function() {
  testUtil.assertEquals([2, 4, 6].toString(), [1, 2, 3].each(function(x) {return x * 2;}).toString());
}

Test.Array.testInclude = function() {
  testUtil.assertEquals(true, [2, 4, 6].include(2));
}

Test.Array.testInclude2 = function() {
  testUtil.assertEquals(false, ["a", "b", "c"].include("d"));
}
