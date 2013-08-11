TestCase("HTTP Test", {
  testGet: function() {
    try {
      var result = HTTP.get("www.yahoo.co.jp", "/index.html", 80);
     assert(true);
    } catch (e) {
      assert(false);
    }
  },
  testGet2: function() {
    try {
      var result = HTTP.get("www.yahoo.co.jp", "/indexa.html", 80);
      assert(false);
    } catch (e) {
      assert(true);
    }
  },
  testGet3: function() {
    try {
      var result = HTTP.get("www.yahooooo.co.jp", "/index.html", 80);
      assert(false);
    } catch (e) {
      assert(true);
    }
  },
  testStart: function() {
    try {
      HTTP.start("www.yahoo.co.jp", 80, function(http) {
        var res = http.get("/index.html");
        assert(true);
      });
    } catch (e) {
      assert(false);
    }
  },
  testStartWithHeader: function() {
    try {
      HTTP.start("www.yahoo.co.jp", 80, function(http) {
        var res = http.get("/index.html", {"Connection": "Keep-Alive"});
        assert(true);
      });
    } catch(e) {
      assert(false);
    }
  }
});
