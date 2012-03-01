HTTP.start("www.yahoo.co.jp", 80, function(http) {
  print(http.get("/index.html"));
});

