Test.SpecialFolders = {};

Test.SpecialFolders.allUserProfile = getEnv("ALLUSERSPROFILE");
Test.SpecialFolders.userProfile = getEnv("USERPROFILE");
Test.SpecialFolders.systemRoot = getEnv("SystemRoot");

Test.SpecialFolders.testGetAllUsersDesktop = function() {
  testUtil.assertEquals(Test.SpecialFolders.allUserProfile + "\\デスクトップ", SpecialFolders.getAllUsersDesktop());
}

Test.SpecialFolders.testGetAllUsersPrograms = function() {
  testUtil.assertEquals(Test.SpecialFolders.allUserProfile + "\\スタート メニュー\\プログラム", SpecialFolders.getAllUsersPrograms());
}

Test.SpecialFolders.testGetAllUsersStartMenu = function() {
  testUtil.assertEquals(Test.SpecialFolders.allUserProfile + "\\スタート メニュー", SpecialFolders.getAllUsersStartMenu());
}

Test.SpecialFolders.testGetAllUsersStartup = function() {
  testUtil.assertEquals(Test.SpecialFolders.allUserProfile + "\\スタート メニュー\\プログラム\\スタートアップ", SpecialFolders.getAllUsersStartup());
}

Test.SpecialFolders.testGetDesktop = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\デスクトップ", SpecialFolders.getDesktop());
}

Test.SpecialFolders.testGetFavorites = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\Favorites", SpecialFolders.getFavorites());
}

Test.SpecialFolders.testGetFonts = function() {
  testUtil.assertEquals(Test.SpecialFolders.systemRoot + "\\Fonts", SpecialFolders.getFonts());
}

Test.SpecialFolders.testGetMyDocuments = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\My Documents", SpecialFolders.getMyDocuments());
}

Test.SpecialFolders.testGetNetHood = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\VSWebCache\\NetHood", SpecialFolders.getNetHood());
}

Test.SpecialFolders.testGetPrintHood = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\PrintHood", SpecialFolders.getPrintHood());
}

Test.SpecialFolders.testGetPrograms = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\スタート メニュー\\プログラム", SpecialFolders.getPrograms());
}

Test.SpecialFolders.testGetRecent = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\Recent", SpecialFolders.getRecent());
}

Test.SpecialFolders.testGetSendTo = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\SendTo", SpecialFolders.getSendTo());
}

Test.SpecialFolders.testGetStartMenu = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\スタート メニュー", SpecialFolders.getStartMenu());
}

Test.SpecialFolders.testGetStartup = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\スタート メニュー\\プログラム\\スタートアップ", SpecialFolders.getStartup());
}

Test.SpecialFolders.testGetTemplates = function() {
  testUtil.assertEquals(Test.SpecialFolders.userProfile + "\\Templates", SpecialFolders.getTemplates());
}