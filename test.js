/*
File.open("hoge.txt", "w", function(outfile) {
  File.open("mado.js", "r", function(infile) {
    infile.each(function(line) {
      outfile.puts(line);
    });
  });
});
*/

//WScript.Echo(File.size("hage"));

//WScript.Echo(Dir.pwd());

/*
Dir.cd("old", function(path) {
  File.open("hoge.txt", "w", function(outfile) {
    outfile.puts("adsfsadflsadjf");
  });
});

File.open("moge.txt", "w", function(outfile) {
  outfile.puts("sdfasdfasdadsfsadflsadjf");
});

*/


//WScript.Echo(File.realpath("test.js", "old"));

//Dir.create("hage");

/*
print("hoge");
sleep(2);
print("moge");
*/

//exit(100);


//WScript.Echo("   	abc			".trim());

/*
print(new Date().getYYYYMMDD());
print(new Date().getYYYYMMDDHH24MISS());
print(new Date().getFormattedDate());
*/
/*
var x = [1, 2, 3].each(function(i) {
  return i*2;
});

print(x);
*/

//print(getComputerName());

//Process.exec("notepad", ["test.js"], Mado.WINDOW_STYLE_NORMAL, false);

//print("abc");

/*
if(isWScriptRunning()) {
  print("true");
} else {
  print("false");
}
*/

/*
print(getEnv("COMPUTERNAME"));

sendKeys("abc");
*/

/*
var c = new Clipboard();
c.set("“ú–{Œê");
sendKeys("^v");
c.close();
*/

/*
osShutDown(10);
osReboot(10);
*/

/*
print(File.extname("c:\work"));
print(File.extname("c:\work\build_y2.xml"));
*/

/*
print(Dir.getParentDir("c:\\work"));
print(Dir.getParentDir("c:\\work\\build_y2.xml"));

var d = new Dir("c:\\app\\adblockfeed\\static");
print(d.getParentDir());

*/
/*
HTTP.start("www.yahoo.co.jp", 80, function(http) {
  print(http.get("/index.html"));
});
*/
/*
print(HTTP.get("www.yahoo.co.jp", "/index.html", 80));
HTTP.get_print("www.yahoo.co.jp", "/index.html", 80);
*/

/*
File.open("aaa.txt", "a", function(outfile) {
  outfile.puts("hage");
  outfile.puts("moge");
  outfile.puts("toge");
});
*/

/*
var d = new Dir("c:\\work\\diff");
print(d.getParentDir());

var d2 = Dir.open("c:\\work\\diff");
print(d2.getParentDir());

Dir.open("c:\\work\\diff", function(d3) {
  print(d3.getParentDir());
});
*/

/*
var d = new Dir("c:\\work\\diff");
d.each(function(item) {
  print(item);
});
*/

/*
Dir.entries("c:\\work\\diff").each(function(item) {
  print(item);
});
*/
/*
Dir.find("c:\\work\\diff").each(function(item) {
  print(item);
});
*/

Dir.find("c:\\work\\diff", "java.*.bat", function(item) {
  print(item);
});


//print(File.basename("c:\\work\\diff\\svndiff_java.bat"));