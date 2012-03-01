// Global Variables
/**
 * 引数で構成された配列
 */
var ARGV = [];
(function() {
  var argvLength = WScript.Arguments.length;
  for(var i=0; i<argvLength; i++) {
    ARGV.push(WScript.Arguments(i));
  }
}());

/**
 * スクリプトのフルパス
 */
var __FILE__ = WScript.ScriptFullName;

// Global Functions
/**
 * 指定したmsec(ミリ秒)スリープします。
 * @param {Number} msec スリープするミリ秒
 */
var sleep = function(msec) {
  WScript.Sleep(msec);
};

/**
 * 指定したfnがtrueを返すまで、msec(ミリ秒)間隔でfnを実行します。
 * @param {Number} msec スリープするミリ秒
 * @param {Function} fn 実行する関数
 * @param {Array} args 関数に渡す引数
 */
var sleepif = function(msec, fn, args) {
  while(true) {
    if(fn(args)) {
      return true;
    }
    sleep(msec);
  }
};

/**
 * msgを表示します。
 * @param {String} msg メッセージ
 */
var print = function(msg) {
  WScript.Echo(msg);
};

/**
 * プログラムを指定したステータスで終了します。
 * @param {Number} status (オプション)終了ステータス(初期値 0)
 */
var exit = function(status) {
  var exitStatus = status || 0;
  WScript.Quit(exitStatus);
};

/**
 * fnが関数かどうか判定します。
 * @param {Object} fn オブジェクト
 * @return {Boolean} 関数の場合はtrue、そうでない場合はfalseを返す。
 */
var isFunction = function (fn) {
  return (typeof fn == "function");
};

/**
 * ホスト名を取得します。
 * @return {String} ホスト名
 */
var getComputerName = function () {
  return Const.WSHELL.ExpandEnvironmentStrings("%COMPUTERNAME%");
};

/**
 * WScript.exeで実行しているかどうか判定します。
 * @return {Boolean} WScript.exeで実行している場合はtrue、そうでない場合(CScript.exe)はfalseを返す。
 */
var isWScriptRunning = function() {
  return /wscript\.exe$/i.test(WScript.FullName);
};

/**
 * 指定したkeyのキー入力を行います。
 * @param {String} key 入力する文字列
 * @param {Number} num (オプション)送信する回数(初期値 1)
 * @param {Number} duration (オプション)送信までの待ち時間(ミリ秒)(初期値 0)
 */
var sendKeys = function (key, num, duration) {
  var myNum = num || 1;
  var myDuration = duration || 0;
  for(var i=0; i<myNum; i++) {
    sleep(myDuration);
    Const.WSHELL.Sendkeys(key);
  }
  return this;
};

/**
 * 指定したnameの環境変数を取得します。
 * @param {String} name 取得する環境変数名
 * @return {String} 環境変数の値
 */
var getEnv = function (name) {
  return Const.WSHELL.ExpandEnvironmentStrings("%" + name + "%");
};

/**
 * OSをシャットダウンします。
 * @param {Number} timeout (オプション)タイムアウトを秒で指定(初期値 0)
 */
var osShutdown = function (timeout) {
  var myTimeout = timeout || 0;
  Process.exec("shutdown", ["-s", "-t", myTimeout], 0, false);
};

/**
 * OSを再起動します。
 * @param {Number} timeout (オプション)タイムアウトを秒で指定(初期値 0)
 * @param {String} option "-f"を指定すると、実行中のプロセスを警告なしで閉じます
 */
var osReboot = function (timeout, option) {
  var params = ["-r", "-t"];
  var myTimeout = timeout || 0;
  params.push(myTimeout);

  if (option !== null) {
    params.push(option);
  }

  Process.exec("shutdown", params, 0, false);
};

/**
 * ショートカットを作成します。
 * @param {String} pathTo 作成するショートカットの場所
 * @param {String} pathFrom ショートカットの対象となるファイル、または、URL
 * @example 使用例：
// SendToにノートパッドのショートカットを作成
createShortcut(SpecialFolders.getSendTo() + "\\notepad.lnk", "notepad.exe");

// デスクトップにYahooのURLショートカットを作成
createShortcut(SpecialFolders.getDesktop() + "\\yahoo.url", "http://www.yahoo.co.jp/");
 */
var createShortcut = function(pathTo, pathFrom) {
  var shortcut = Const.WSHELL.CreateShortcut(pathTo);
  shortcut.TargetPath = pathFrom;
  shortcut.Save();
};

/**
 * ポップアップメッセージを表示します。
 * @param {String} strText 表示するメッセージ
 * @param {Number} nSecondsToWait (オプション)ポップアップウインドウを閉じるまでの秒数(初期値 0)
 * @param {String} strTitle (オプション)ポップアップウインドウのタイトル(初期値 "Windows Script Host")
 * @param {Number} buttonType (オプション)ポップアップウインドウのボタンの種類(初期値 Const.BUTTON_TYPE_OK)
 * @param {Number} iconType (オプション)ポップアップウインドウのアイコンの種類(初期値 Const.ICON_TYPE_INFO)
 * @return {Number} 選択されたボタン値
 * @example 使用例：
if(popup("mado.jsを使いますか？"
   , 0
   , "title"
   , Const.BUTTON_TYPE_YES_NO) === Const.BUTTON_VALUE_YES) {
   print("はい！");
}
 */
var popup = function(strText, nSecondsToWait, strTitle, buttonType, iconType) {
  var myNSecondsToWait = nSecondsToWait || 0;
  var myStrTitle = strTitle || "Windows Script Host";
  var myButtonType = buttonType || Const.BUTTON_TYPE_OK_ONLY;
  var myIconType = iconType || Const.ICON_TYPE_INFO;
  return Const.WSHELL.popup(strText, myNSecondsToWait, myStrTitle, myButtonType + myIconType);
};

// ----------------------------------------
// Import VBScript Source Code
// ----------------------------------------
var objJS = new ActiveXObject("ScriptControl");
objJS.Language = "VBScript";
objJS.AddCode(
'Function JSInputBox(prompt, title) ' + 
'  JSInputBox = InputBox(prompt, title) ' + 
'End Function'
);
// ----------------------------------------

/**
 * インプットボックスを表示します。
 * @param {String} prompt (オプション)ダイアログボックス内にメッセージとして表示する文字列を示す文字列式を指定します。(初期値 空文字)
 * @param {String} title (オプション)ダイアログボックスのタイトルバーに表示する文字列を示す文字列式を指定します。(初期値 空文字)
 */
var inputBox = function(prompt, title) {
  var myPrompt = prompt || "";
  var myTitle = title || "";

  return objJS.CodeObject.JSInputBox(myPrompt, myTitle);
};

/**
 * JScriptの一次元配列を、一次元のSafeArrayに変換する。<br/>
 * http://www.imasy.or.jp/~hir/hir/tech/js_tips.html#safearray を参考にしました。
 * @param {Array} jsArray JScriptの配列を指定します。
 * @return {Object} 変換されたSafeArrayを返します。通常、VBArrayでラップして使います。
 */
var arrayToSafeArray = function(jsArray) {
  var dic = new ActiveXObject("Scripting.Dictionary");
  var jsArrayLength = jsArray.length;
  for (var i = 0; i < jsArrayLength; i++) {
    dic.add(i, jsArray[i]);
  }

  return dic.items();
}