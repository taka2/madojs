Test.Registry = {};

Test.Registry.testRegRead = function() {
  try {
    var val = Registry.regRead("HKCU\\Software\\Microsoft\\Windows Script Host\\");
    testUtil.assertTrue(true);
  } catch(e) {
    testUtil.assertTrue(false);
  }
};
