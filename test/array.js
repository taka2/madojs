Test.Array = {};

Test.Array.testEach = function() {
  testUtil.assertEquals([2, 4, 6].toString(), [1, 2, 3].each(function(x) {return x * 2;}).toString());
}
