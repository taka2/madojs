TestCase("Registry Test", {
  testRegRead: function() {
    try {
      var val = Registry.regRead("HKCU\\Software\\Microsoft\\Windows Script Host\\");
      assert(true);
    } catch(e) {
      assert(false);
    }
  }
});
