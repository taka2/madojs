var path = "/projects/ruby-trunk/issues.csv?c%5B%5D=tracker&c%5B%5D=status&c%5B%5D=priority&c%5B%5D=subject&c%5B%5D=assigned_to&c%5B%5D=updated_on&f%5B%5D=status_id&f%5B%5D=priority_id&f%5B%5D=&group_by=&op%5Bpriority_id%5D=%3D&op%5Bstatus_id%5D=%2A&set_filter=1&v%5Bpriority_id%5D%5B%5D=3";

HTTP.start("bugs.ruby-lang.org", 80, function(http) {
  File.open("redmine.csv", "w", function(file) {
    var response = http.get(path, {"Accept-Encoding": "Shift_JIS"}, HTTP.RESPONSE_TYPE_BODY);
    file.puts(Stream.binaryToText(response, "Shift_JIS"));
  });
});

print("done");
