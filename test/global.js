TestCase("Global Test", {
  testSleep: function() {
    sleep(500);
    assert(true);
  },
  testSleepIf: function() {
    sleepif(500, function() {return true;});
    assert(true);

    sleepif(500, function(a) {if(a>10) {return true;}}, [11]);
    assert(true);
  },
  testIsFunction: function() {
    assert(!isFunction(1.23));
    assert(!isFunction("abc"));
    assert(!isFunction({}));
    assert(!isFunction(new Date()));
    assert(isFunction(function() {return true;}));
  },
  testGetComputerName: function() {
    var computerName = getComputerName();
    assert(computerName);
  },
  testCreateShortcut: function() {
    createShortcut(SpecialFolders.getSendTo() + "\\notepad.lnk", "notepad.exe");
    assert(File.exist(SpecialFolders.getSendTo() + "\\notepad.lnk"));
    File.unlink(SpecialFolders.getSendTo() + "\\notepad.lnk");
  },
  testArrayToSafeArray: function() {
    var testArray = [1,2,3];
    var safeArray = arrayToSafeArray(testArray);

    var wrappedSafeArray = new VBArray(safeArray);
    assertEquals(1, wrappedSafeArray.dimensions());
    assertEquals(0, wrappedSafeArray.lbound());
    assertEquals(2, wrappedSafeArray.ubound());

    for(var i=0; i<testArray.length; i++) {
      assertEquals(testArray[i], wrappedSafeArray.getItem(i));
    }

    var jsArray = wrappedSafeArray.toArray();
    assertEquals(3, jsArray.length);
  }
});
