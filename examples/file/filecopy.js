// filecopy.jsを1行ずつ読みながら、copy_of_filecopy.jsにコピーする。

File.open("copy_of_filecopy.js", "w", function(outfile) {
  File.open("filecopy.js", "r", function(infile) {
    infile.each(function(line) {
      outfile.puts(line);
    });
  });
});
