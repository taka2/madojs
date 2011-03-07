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
