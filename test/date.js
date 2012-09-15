TestCase("Date Test", {
  setUp: function() {
    this.dt = new Date(2011, 2, 22, 15, 25, 28, 3);
    this.dt2 = new Date("2012/10/22 21:30:24");
  },
  testGetFormattedDate: function() {
    assertEquals("2011/03/22 15:25:28", this.dt.getFormattedDate());
  },
  testGetYYYYMMDD: function() {
    assertEquals("20110322", this.dt.getYYYYMMDD());
  },
  testGetYYYYMMDD_2: function() {
    assertEquals("20121022", this.dt2.getYYYYMMDD());
  },
  testGetYYYYMMDDHH24MISS: function() {
    assertEquals("20110322152528", this.dt.getYYYYMMDDHH24MISS());
  },
  testGetYYYYMMDDHH24MISS: function() {
    assertEquals("20121022213024", this.dt2.getYYYYMMDDHH24MISS());
  },
  testAdd: function() {
    // 10•ªŒã
    assertEquals("20121022214024", this.dt2.add(10 * 60 * 1000).getYYYYMMDDHH24MISS());
    // 10•ª‘O
    assertEquals("20121022213024", this.dt2.add(-10 * 60 * 1000).getYYYYMMDDHH24MISS());
  }
});
