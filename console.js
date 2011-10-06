while(true) {
  try {
    WScript.StdOut.Write(">");
    var line = WScript.StdIn.ReadLine();
    print(eval(line));
  } catch (e) {
    print(e.number + ":" + e.description);
  }
}
