Test.Global = {};

Test.Global.testSleep = function() {
  sleep(500);
  testUtil.assertTrue(true);
};

Test.Global.testSleepIf = function() {
  sleepif(500, function() {return true;});
  testUtil.assertTrue(true);

  sleepif(500, function(a) {if(a>10) {return true;}}, [11]);
  testUtil.assertTrue(true);
};

Test.Global.testIsFunction = function() {
  testUtil.assertTrue(!isFunction(1.23));
  testUtil.assertTrue(!isFunction("abc"));
  testUtil.assertTrue(!isFunction({}));
  testUtil.assertTrue(!isFunction(new Date()));
  testUtil.assertTrue(isFunction(function() {return true;}));
};

Test.Global.testGetComputerName = function() {
  var computerName = getComputerName();
  testUtil.assertTrue(computerName);
};

Test.Global.testCreateShortcut = function() {
  createShortcut(SpecialFolders.getSendTo() + "\\notepad.lnk", "notepad.exe");
  testUtil.assertTrue(File.exist(SpecialFolders.getSendTo() + "\\notepad.lnk"));
  File.unlink(SpecialFolders.getSendTo() + "\\notepad.lnk");
};

Test.Global.testArrayToSafeArray = function() {
  var testArray = [1,2,3];
  var safeArray = arrayToSafeArray(testArray);

  var wrappedSafeArray = new VBArray(safeArray);
  testUtil.assertEquals(1, wrappedSafeArray.dimensions());
  testUtil.assertEquals(0, wrappedSafeArray.lbound());
  testUtil.assertEquals(2, wrappedSafeArray.ubound());

  for(var i=0; i<testArray.length; i++) {
    testUtil.assertEquals(testArray[i], wrappedSafeArray.getItem(i));
  }

  var jsArray = wrappedSafeArray.toArray();
  testUtil.assertEquals(3, jsArray.length);
};