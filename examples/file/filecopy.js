// filecopy.js��1�s���ǂ݂Ȃ���Acopy_of_filecopy.js�ɃR�s�[����B

File.open("copy_of_filecopy.js", "w", function(outfile) {
  File.open("filecopy.js", "r", function(infile) {
    infile.each(function(line) {
      outfile.puts(line);
    });
  });
});
