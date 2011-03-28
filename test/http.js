Test.HTTP = {};

Test.HTTP.testGet = function() {
  try {
    var result = HTTP.get("www.yahoo.co.jp", "/index.html", 80);
    testUtil.assertTrue(true);
  } catch (e) {
    testUtil.assertTrue(false);
  }
}

Test.HTTP.testGet2 = function() {
  try {
    var result = HTTP.get("www.yahoo.co.jp", "/indexa.html", 80);
    testUtil.assertTrue(false);
  } catch (e) {
    testUtil.assertTrue(true);
  }
}

Test.HTTP.testGet3 = function() {
  try {
    var result = HTTP.get("www.yahooooo.co.jp", "/index.html", 80);
    testUtil.assertTrue(false);
  } catch (e) {
    testUtil.assertTrue(true);
  }
}

Test.HTTP.testStart = function() {
  try {
    HTTP.start("www.yahoo.co.jp", 80, function(http) {
      var res = http.get("/index.html");
      testUtil.assertTrue(true);
    });
  } catch (e) {
    testUtil.assertTrue(false);
  }
}
