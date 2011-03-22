Test.Date = {};

Test.Date.dt = new Date(2011, 2, 22, 15, 25, 28, 3);

Test.Date.testGetFormattedDate = function() {
  testUtil.assertEquals("2011/03/22 15:25:28", Test.Date.dt.getFormattedDate());
}

Test.Date.testGetYYYYMMDD = function() {
  testUtil.assertEquals("20110322", Test.Date.dt.getYYYYMMDD());
}

Test.Date.testGetYYYYMMDDHH24MISS = function() {
  testUtil.assertEquals("20110322152528", Test.Date.dt.getYYYYMMDDHH24MISS());
}
