Test.Date = {};

Test.Date.dt = new Date(2011, 2, 22, 15, 25, 28, 3);

Test.Date.testGetFormattedDate = function() {
  testUtil.assertEquals(Test.Date.dt.getFormattedDate(), "2011/03/22 15:25:28");
}

Test.Date.testGetYYYYMMDD = function() {
  testUtil.assertEquals(Test.Date.dt.getYYYYMMDD(), "20110322");
}

Test.Date.testGetYYYYMMDDHH24MISS = function() {
  testUtil.assertEquals(Test.Date.dt.getYYYYMMDDHH24MISS(), "20110322152528");
}