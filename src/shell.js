/**
 * インスタンス化しません。
 * @class シェル(explorer)へのアクセスを提供するクラス
<pre class = "code">
使用例：
// Windowsフォルダをエクスプローラで開く
Shell.open("C:\\windows");
</pre>
 */
var Shell = {};

/**
 * 指定したpathをexplorerで開きます。
 * @param {String} path explorerで開くパス
 */
Shell.open = function (path) {
  Const.SHELLAPP.Open(path);
};

/**
 * 指定したserviceNameのサービスを開始します。
 * @param {String} serviceName 開始するサービス名
 * @param {Boolean} isPersistent (オプション)スタートアップの種類を自動に変更します(初期値 false)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返します。
 */
Shell.serviceStart = function(serviceName, isPersistent) {
  myIsPersistent = isPersistent || false;
  return Const.SHELLAPP.ServiceStart(serviceName, myIsPersistent);
};

/**
 * 指定したserviceNameのサービスを停止します。
 * @param {String} serviceName 停止するサービス名
 * @param {Boolean} isPersistent (オプション)スタートアップの種類を手動に変更します(初期値 false)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返します。
 */
Shell.serviceStop = function(serviceName, isPersistent) {
  myIsPersistent = isPersistent || false;
  return Const.SHELLAPP.ServiceStop(serviceName, myIsPersistent);
};
