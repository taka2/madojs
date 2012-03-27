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

/**
 * 指定したzipFileを、指定したフォルダに解凍します。
 * @param {String} zipFile 解凍するzipファイル(フルパス)
 * @param {String} unzipTo 解凍先フォルダ
 * @throws zipFile、または、unzipToディレクトリが存在しない場合にスローされます。
 */
Shell.unzip = function(zipFile, unzipTo) {
  if(!File.exist(zipFile)) {
    throw new Error(-1, "File Not Found: " + zipFile);
  }
  if(!Dir.exist(unzipTo)) {
    throw new Error(-1, "Path Not Found: " + unzipTo);
  }

  var zipFileObj = Const.SHELLAPP.Namespace(zipFile);
  var zipFileItems = zipFileObj.Items();
  var unzipToObj = Const.SHELLAPP.Namespace(unzipTo);

  var i;
  for(i=0; i<zipFileItems.Count; i++) {
    unzipToObj.copyhere(zipFileItems.item(i));
  }
};

/**
 * 指定したフォルダを、zipFileに圧縮します。
 * @param {String} zipFileTo 作成するzipファイル(フルパス)
 * @param {String} targetFolder 圧縮対象フォルダ
 * @throws targetFolderディレクトリが存在しない場合にスローされます。
 */
Shell.zip = function(zipFileTo, targetFolder) {
  if(!Dir.exist(targetFolder)) {
    throw new Error(-1, "Path Not Found: " + unzipTo);
  }

  File.open(zipFileTo, "w", function(file) {
    file.write("PK\5\6\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0");
  });

  var zipFileTo = Const.SHELLAPP.Namespace(zipFileTo);
  var targetDirectory = Const.SHELLAPP.Namespace(targetFolder);
  var targetDirectoryItems = targetDirectory.Items();

  var i;
  for(i=0; i<targetDirectoryItems.Count; i++) {
    var o = targetDirectoryItems.item(i);
    zipFileTo.CopyHere(o);
  }

  while(true) {
    sleep(100);
    if(zipFileTo.Items().Count === targetDirectoryItems.Count) {
      break;
    }
  }
};
