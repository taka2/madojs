// Access mdb�ɑ΂��āA�w�肵��SQL�t�@�C���̓��e���ꊇ�Ŏ��s����B
if(ARGV.length != 2) {
  print("Usage: ExecuteSQLFileForMdb.wsf [mdbFileName] [sqlFileName]");
  exit(1);
}
var mdbFileName = ARGV[0];
var sqlFileName = ARGV[1];

AdoAccessConnection.open(mdbFileName, "", "", function(con) {
  File.open(sqlFileName, "r", function(infile) {
    infile.each(function(line) {
      con.executeUpdate(line);
    });
  });
});
