/**
 * インスタンス化しません。
 * @class 特定のフォルダパスを取得する機能を提供するクラス
<pre class = "code">
使用例：
// デスクトップのパスを取得
print(SpecialFolders.getDesktop());
</pre>
 */
var SpecialFolders = {};

/**
 * AllUsersのデスクトップパスを取得します。
 * @return {String} AllUsersのデスクトップパス
 */
SpecialFolders.getAllUsersDesktop = function () {
  return Const.WSHELL.SpecialFolders("AllUsersDesktop");
};

/**
 * AllUsersのスタートメニューパスを取得します。
 * @return {String} AllUsersのスタートメニューパス
 */
SpecialFolders.getAllUsersStartMenu = function () {
  return Const.WSHELL.SpecialFolders("AllUsersStartMenu");
};

/**
 * AllUsersのプログラムパスを取得します。
 * @return {String} AllUsersのプログラムパス
 */
SpecialFolders.getAllUsersPrograms = function () {
  return Const.WSHELL.SpecialFolders("AllUsersPrograms");
};

/**
 * AllUsersのスタートアップパスを取得します。
 * @return {String} AllUsersのスタートアップパス
 */
SpecialFolders.getAllUsersStartup = function () {
  return Const.WSHELL.SpecialFolders("AllUsersStartup");
};

/**
 * デスクトップパスを取得します。
 * @return {String} デスクトップパス
 */
SpecialFolders.getDesktop = function () {
  return Const.WSHELL.SpecialFolders("Desktop");
};

/**
 * お気に入りパスを取得します。
 * @return {String} お気に入りパス
 */
SpecialFolders.getFavorites = function () {
  return Const.WSHELL.SpecialFolders("Favorites");
};

/**
 * フォントパスを取得します。
 * @return {String} フォントパス
 */
SpecialFolders.getFonts = function () {
  return Const.WSHELL.SpecialFolders("Fonts");
};

/**
 * マイドキュメントパスを取得します。
 * @return {String} マイドキュメントパス
 */
SpecialFolders.getMyDocuments = function () {
  return Const.WSHELL.SpecialFolders("MyDocuments");
};

/**
 * ネットフードパスを取得します。
 * @return {String} ネットフードパス
 */
SpecialFolders.getNetHood = function () {
  return Const.WSHELL.SpecialFolders("NetHood");
};

/**
 * プリントフードパスを取得します。
 * @return {String} プリントフードパス
 */
SpecialFolders.getPrintHood = function () {
  return Const.WSHELL.SpecialFolders("PrintHood");
};

/**
 * プログラムパスを取得します。
 * @return {String} プログラムパス
 */
SpecialFolders.getPrograms = function () {
  return Const.WSHELL.SpecialFolders("Programs");
};

/**
 * 最近使ったファイルパスを取得します。
 * @return {String} 最近使ったファイルパス
 */
SpecialFolders.getRecent = function () {
  return Const.WSHELL.SpecialFolders("Recent");
};

/**
 * 送るパスを取得します。
 * @return {String} 送るパス
 */
SpecialFolders.getSendTo = function () {
  return Const.WSHELL.SpecialFolders("SendTo");
};

/**
 * スタートメニューパスを取得します。
 * @return {String} スタートメニューパス
 */
SpecialFolders.getStartMenu = function () {
  return Const.WSHELL.SpecialFolders("StartMenu");
};

/**
 * スタートアップパスを取得します。
 * @return {String} スタートアップパス
 */
SpecialFolders.getStartup = function () {
  return Const.WSHELL.SpecialFolders("Startup");
};

/**
 * テンプレートパスを取得します。
 * @return {String} テンプレートパス
 */
SpecialFolders.getTemplates = function () {
  return Const.WSHELL.SpecialFolders("Templates");
};
