TestCase("Date Test", {
  setUp: function() {
    this.dt = new Date(2011, 2, 22, 15, 25, 28, 3);
  },
  testGetFormattedDate: function() {
    assertEquals("2011/03/22 15:25:28", this.dt.getFormattedDate());
  },
  testGetYYYYMMDD: function() {
    assertEquals("20110322", this.dt.getYYYYMMDD());
  },
  testGetYYYYMMDDHH24MISS: function() {
    assertEquals("20110322152528", this.dt.getYYYYMMDDHH24MISS());
  }
});
