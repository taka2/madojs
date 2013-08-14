/**
 * インスタンス化しません。
 * @class MadoJSの定数を定義するクラス。
 */
var Const = {};

// Static members of Global
Const.FSO = new ActiveXObject("Scripting.FileSystemObject");
Const.WSHELL = new ActiveXObject("WScript.Shell");
Const.SHELLAPP = new ActiveXObject("Shell.Application");

// Static Variables
Const.INITIAL_CURRENT_DIRECTORY = Const.WSHELL.CurrentDirectory;

/**
 * Window Styleの定数：通常(1)
 */
Const.WINDOW_STYLE_NORMAL = 1; // 通常

/**
 * Window Styleの定数：最小化(2)
 */
Const.WINDOW_STYLE_MIN = 2;    // 最小化

/**
 * Window Styleの定数：最大化(3)
 */
Const.WINDOW_STYLE_MAX = 3;    // 最大化

/**
 * IO Modeの定数：読み取り専用(1)
 */
Const.IOMODE_FORREADING = 1;   // 読み取り専用

/**
 * IO Modeの定数：書き込み専用(2)
 */
Const.IOMODE_FORWRITING = 2;   // 書き込み専用

/**
 * IO Modeの定数：追加書き込み(8)
 */
Const.IOMODE_FORAPPENDING = 8; // 追加書き込み

/**
 * LogEventタイプ：SUCCESS(0)
 */
Const.LOG_EVENT_TYPE_SUCCESS = 0;

/**
 * LogEventタイプ：ERROR(1)
 */
Const.LOG_EVENT_TYPE_ERROR = 1;

/**
 * LogEventタイプ：WARNING(2)
 */
Const.LOG_EVENT_TYPE_WARNING = 2;

/**
 * LogEventタイプ：INFORMATION(4)
 */
Const.LOG_EVENT_TYPE_INFORMATION = 4;

/**
 * LogEventタイプ：AUDIT_SUCCESS(8)
 */
Const.LOG_EVENT_TYPE_AUDIT_SUCCESS = 8;

/**
 * LogEventタイプ：AUDIT_FAILURE(16)
 */
Const.LOG_EVENT_TYPE_AUDIT_FAILURE = 16;

/**
 * ボタンタイプ：OK_ONLY(0)
 */
Const.BUTTON_TYPE_OK_ONLY = 0;

/**
 * ボタンタイプ：OK_CANCEL(1)
 */
Const.BUTTON_TYPE_OK_CANCEL = 1;

/**
 * ボタンタイプ：STOP_RETRY_IGNORE(2)
 */
Const.BUTTON_TYPE_STOP_RETRY_IGNORE = 2;

/**
 * ボタンタイプ：YES_NO_CANCEL(3)
 */
Const.BUTTON_TYPE_YES_NO_CANCEL = 3;

/**
 * ボタンタイプ：YES_NO(4)
 */
Const.BUTTON_TYPE_YES_NO = 4;

/**
 * ボタンタイプ：RETRY_CANCEL(5)
 */
Const.BUTTON_TYPE_RETRY_CANCEL = 5;

/**
 * アイコンタイプ：STOP(16)
 */
Const.ICON_TYPE_STOP = 16;

/**
 * アイコンタイプ：QUESTION(32)
 */
Const.ICON_TYPE_QUESTION = 32;

/**
 * アイコンタイプ：EXCLAMATION(48)
 */
Const.ICON_TYPE_EXCLAMATION = 48;

/**
 * アイコンタイプ：INFO(64)
 */
Const.ICON_TYPE_INFO = 64;

/**
 * ボタン値：OK(1)
 */
Const.BUTTON_VALUE_OK = 1;

/**
 * ボタン値：CANCEL(2)
 */
Const.BUTTON_VALUE_CANCEL = 2;

/**
 * ボタン値：STOP(3)
 */
Const.BUTTON_VALUE_STOP = 3;

/**
 * ボタン値：RETRY(4)
 */
Const.BUTTON_VALUE_RETRY = 4;

/**
 * ボタン値：IGNORE(5)
 */
Const.BUTTON_VALUE_IGNORE = 5;

/**
 * ボタン値：YES(6)
 */
Const.BUTTON_VALUE_YES = 6;

/**
 * ボタン値：NO(7)
 */
Const.BUTTON_VALUE_NO = 7;
// Global Variables
/**
 * 引数で構成された配列<br/>
 * hta環境では[]となります。
 */
var ARGV = [];
(function() {
  try {
    var argvLength = WScript.Arguments.length;
    for(var i=0; i<argvLength; i++) {
      ARGV.push(WScript.Arguments(i));
    }
  } catch(e) {
    // WScriptが未定義の場合、ARGV = []
  }
}());

/**
 * スクリプトのフルパス<br/>
 * hta環境ではundefinedとなります。
 */
try {
  var __FILE__ = WScript.ScriptFullName;
} catch(e) {
  if(e.name === 'TypeError') {
    // WScriptが未定義の場合、__FILE__ = undefined
  } else {
    throw e;
  }
}

// Global Functions
/**
 * 指定したmsec(ミリ秒)スリープします。<br/>
 * hta環境では何もしません。
 * @param {Number} msec スリープするミリ秒
 */
var sleep = function(msec) {
  try {
    WScript.Sleep(msec);
  } catch(e) {
    if(e.name === 'TypeError') {
      // 何もしない
    } else {
      throw e;
    }
  }
};

/**
 * 指定したfnがtrueを返すまで、msec(ミリ秒)間隔でfnを実行します。
 * @param {Number} msec スリープするミリ秒
 * @param {Function} fn 実行する関数
 * @param {Array} args (オプション)関数に渡す引数
 */
var sleepif = function(msec, fn, args) {
  if(!args) {
    args = [];
  }
  while(true) {
    if(fn.apply({}, args)) {
      return true;
    }

    sleep(msec);
  }
};

/**
 * msgを表示します。<br/>
 * hta環境では印刷ダイアログが表示されるため、代わりに<a href = "_global_.html#echo">echo</a>を使用してください。
 * @param {String} msg メッセージ
 */
var print = function(msg) {
  try {
    WScript.Echo(msg);
  } catch(e) {
    if(e.name === 'TypeError') {
      window.alert(msg);
    } else {
      throw e;
    }
  }
};

/**
 * msgを表示します。<br/>
 * hta環境ではwindow.alertを実行します。
 * 
 * @param {String} msg メッセージ
 */
var echo = print;

/**
 * プログラムを指定したステータスで終了します。<br/>
 * hta環境ではwindow.closeを実行します。
 * @param {Number} status (オプション)終了ステータス(初期値 0)
 */
var exit = function(status) {
  var exitStatus = status || 0;
  try {
    WScript.Quit(exitStatus);
  } catch(e) {
    if(e.name === 'TypeError') {
      window.close();
    } else {
      throw e;
    }
  }
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
 * WScript.exeで実行しているかどうか判定します。<br/>
 * hta環境の場合は常にfalseを返します。
 * @return {Boolean} WScript.exeで実行している場合はtrue、そうでない場合(CScript.exe)はfalseを返す。
 */
var isWScriptRunning = function() {
  try {
    return /wscript\.exe$/i.test(WScript.FullName);
  } catch(e) {
    if(e.name === 'TypeError') {
      return false;
    } else {
      throw e;
    }
  }
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
try {
  var objJS = new ActiveXObject("ScriptControl");
  objJS.Language = "VBScript";
  objJS.AddCode(
  'Function JSInputBox(prompt, title) ' + 
  '  JSInputBox = InputBox(prompt, title) ' + 
  'End Function'
  );
} catch(e) {
  // 何もしない
}
// ----------------------------------------

/**
 * インプットボックスを表示します。
 * @param {String} prompt (オプション)ダイアログボックス内にメッセージとして表示する文字列を示す文字列式を指定します。(初期値 空文字)
 * @param {String} title (オプション)ダイアログボックスのタイトルバーに表示する文字列を示す文字列式を指定します。(初期値 空文字)
 */
var inputBox = function(prompt, title) {
  var myPrompt = prompt || "";
  var myTitle = title || "";

  if(objJS) {
    return objJS.CodeObject.JSInputBox(myPrompt, myTitle);
  }
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
}/**
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
// オリジナルのtoStringをorgToStringとして保存
Array.prototype.orgToString = Array.prototype.toString;

/**
 * オブジェクトの値を表す文字列を返します。
 * @return {String} オブジェクトの値を表す文字列
 */
Array.prototype.toString = function() {
  var result = "[";
  var length = this.length;
  for(var i=0; i<length; i++) {
    if(result !== "[") {
      result += ", ";
    }
    result += this[i].toString();
  }
  result += "]";
  return result;
};

/** 
 * 各要素に対してブロックを評価します。
 * @param {Function} block ブロック
 * @example 使用例：
 * var x = [1, 2, 3].forEach(function(i) {
 *   return i*2;
 * });
 *
 * print(x); // 2,4,6
 */
Array.prototype.forEach = function(block) {
  var result = [];

  var thisLength = this.length;
  for(var i=0; i<thisLength; i++) {
    result.push(block(this[i]));
  }

  return result;
};

Array.prototype.each = Array.prototype.forEach;

/**
 * objが配列に含まれるかどうかチェックします。
 * @param {Object} obj 配列に含まれるかどうかチェックするオブジェクト
 * @return {Boolean} objが配列に含まれる場合true、そうでない場合falseを返す。
 */
Array.prototype.include = function(obj) {
  var thisLength = this.length;
  for(var i=0; i<thisLength; i++) {
    if(this[i] === obj) {
      return true;
    }
  }
  return false;
};
/** 
 * 日付をYYYYMMDD形式で返します。
 * @return {String} YYYYMMDD形式の日付
 * @example 使用例：
print(new Date().getYYYYMMDD());
 */
Date.prototype.getYYYYMMDD = function() {
  var yyyy = this.getFullYear();
  var mm = this.getMonth() + 1;
  var dd = this.getDate();
  if (mm < 10) mm = "0" + mm;
  if (dd < 10) dd = "0" + dd;
  return "" + yyyy + mm + dd;
};

/** 
 * 日付をYYYYMMDDHH24MISS形式で返します。
 * @return {String} YYYYMMDDHH24MISS形式の日付
 * @example 使用例：
print(new Date().getYYYYMMDDHH24MISS());
 */
Date.prototype.getYYYYMMDDHH24MISS = function() {
  var hh = this.getHours();
  var mi = this.getMinutes();
  var ss = this.getSeconds();
  if (hh < 10) hh = "0" + hh;
  if (mi < 10) mi = "0" + mi;
  if (ss < 10) ss = "0" + ss;
  return this.getYYYYMMDD() + hh + mi + ss;
};

/** 
 * 日付をYYYY/MM/DD HH:MI:SS形式で返します。
 * @return {String} YYYY/MM/DD HH:MM:SS形式の日付
 * @example 使用例：
print(new Date().getFormattedDate());
 */
Date.prototype.getFormattedDate = function() {
  var yyyy = this.getFullYear();
  var mm = this.getMonth() + 1;
  var dd = this.getDate();
  var hh = this.getHours();
  var mi = this.getMinutes();
  var ss = this.getSeconds();
  if (mm < 10) mm = "0" + mm;
  if (dd < 10) dd = "0" + dd;
  if (hh < 10) hh = "0" + hh;
  if (mi < 10) mi = "0" + mi;
  if (ss < 10) ss = "0" + ss;
  return yyyy + "/" + mm + "/" + dd + " " + hh + ":" + mi + ":" + ss;
};

/**
 * 日付に指定ミリ秒足します。
 * @param {Number} msec 日付に足すミリ秒
 * @return {Date} 指定ミリ秒足した結果の日付
 * @example 使用例：
print(new Date().getFormattedDate());
 */
Date.prototype.add = function(msec) {
  this.setTime(this.getTime() + msec);
  return this;
};/** 
 * オブジェクトの値を表す文字列を返します。
 * @return {String} オブジェクトの値を表す文字列
 */
Object.prototype.toString = function() {
  var result = "{";
  for(e in this) {
    if(this.hasOwnProperty(e)) {
      if(result !== "{") {
        result += ", ";
      }
      result += '"' + e + '": ' + this[e].toString();
    }
  }
  result += "}";
  return result;
};

/**
 * pをプロトタイプに持つオブジェクトを生成します。
 * @param {Object} p プロトタイプ
 * @return {Object} pをプロトタイプに持つオブジェクト
 */
Object.prototype.create = function(p) {
  function f() {};
  f.prototype = p;
  return new f();
};
// オリジナルのtoStringをorgToStringとして保存
String.prototype.orgToString = String.prototype.toString;

/**
 * オブジェクトの値を表す文字列を返します。
 * @return {String} オブジェクトの値を表す文字列
 */
String.prototype.toString = function() {
  return '"' + this.valueOf() + '"';
};

/** 
 * 文字列先頭と末尾のスペースをトリムした文字列を生成して返します。
 * @return {String} トリムした文字列
 * @example 使用例：
print("@" + "   	abc			".trim() + "@"); // "@abc@"
 */
String.prototype.trim = function() {
  var text = this.valueOf();
  return (text || "").replace( /^\s+|\s+$/g, "" );
};

/**
 * 文字列がstrで始まるかどうか
 * @param {String} str テストする文字列
 * @return {Boolean} 文字列がstrで始まる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".startsWith("a"); // true
 */
String.prototype.startsWith = function(str) {
  var text = this.valueOf();
  return(text.match("^" + str) === null ? false : true);
};

/**
 * 文字列がstrで終わるかどうか
 * @param {String} str テストする文字列
 * @return {Boolean} 文字列がstrで終わる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".endsWith("c"); // true
 */
String.prototype.endsWith = function(str) {
  var text = this.valueOf();
  return(text.match(str + "$") === null ? false : true);
};

/**
 * chが全角かどうか判定する。<br/>
 * 二文字以上の文字列でも、先頭の文字のみ判定します。
 * @return {Boolean} chが全角の場合true、そうでない場合falseを返す。
 */
String.prototype.isZenkaku = function() {
  // 先頭一文字のみチェックする
  var text = this.valueOf();
  var c = text.substring(0, 1);

  // escapeした結果%u始まり、かつ、半角カナでない場合全角と判定する
  return (escape(c).substring(0, 2) === "%u") &&
          !c.isHankakuKana();
}

/**
 * chが半角カナかどうか判定する。<br/>
 * 二文字以上の文字列でも、先頭の文字のみ判定します。
 * @return {Boolean} chが半角カナの場合true、そうでない場合falseを返す。
 */
String.prototype.isHankakuKana = function() {
  // 先頭一文字のみチェックする
  var text = this.valueOf();
  var c = text.substring(0, 1);
  var code = c.charCodeAt(0);

  return(code >= 0xFF71 && code <= 0xFF90);
}

/**
 * 全角文字を2バイト、半角文字を1バイトとして、指定位置の文字列を返します。
 * @param {Number} start 開始位置
 * @param {Number} end 終了位置
 * @return {String} 指定位置の文字列
 */
String.prototype.substringb = function(start, end) {
  var text = this.valueOf();
  var currentIndex = 0;
  var result = "";

  var textLength = text.length;
  for(var i=0; i<textLength; i++) {
    var ch = text.charAt(i);
    var chSize = 1;
    if(ch.isZenkaku()) {
      chSize = 2;
    }
    if(currentIndex >= start && (currentIndex + chSize) <= end) {
      result += ch;
    }

    currentIndex += chSize;
  }

  return result;
}

/**
 * 指定文字を、指定文字数に達するまで左側にパディングします。
 * @param {String} ch 文字
 * @param {Number} length 文字数
 * @return {String} 指定位置の文字列
 * @example 使用例：
"12".lpad("0", 5); // "00012"
 */
String.prototype.lpad = function(ch, length) {
  var text = this.valueOf();

  // 文字列が指定文字数に既に達している場合はそのまま返す。
  if(text.length >= length) {
    return text;
  }

  var result = text;
  for(var i=0; i<length - text.length; i++) {
    result = ch + result;
  }
  return result;
}
/** 
 * ファイルオブジェクトを作成します。
 * @class ファイルの操作と読み書きを行うためのクラスです。
<pre class = "code">
使用例：
File.open("copy_of_mado.js", "w", function(outfile) {
  File.open("mado.js", "r", function(infile) {
    infile.each(function(line) {
      outfile.puts(line);
    });
  });
});
</pre>
 * @param {String} path ファイルのパスを文字列で指定します。
 * @param {String} mode ファイルを読み書きするためのモードを指定します。
 * <ul>
 *   <li>'r': 読み取り専用</li>
 *   <li>'w': 書き込み専用</li>
 *   <li>'a': 追加書き込み</li>
 * </ul>
 * @throws モードが'r'または'a'のときで、ファイルが存在しない場合にスローされます。
 */
var File = function(path, mode) {
  this.path = path;
  this.mode = mode;

  if(this.mode === 'w') {
    this.ts = Const.FSO.CreateTextFile(this.path);
  } else if(this.mode === 'a') {
    if(File.exist(this.path)) {
      this.ts = Const.FSO.OpenTextFile(this.path, Const.IOMODE_FORAPPENDING);
    } else {
      this.ts = Const.FSO.CreateTextFile(this.path);
    }
  } else {
    this.ts = Const.FSO.OpenTextFile(this.path, Const.IOMODE_FORREADING);
  }
};

/**
 * パス区切り文字：\
 */
File.PATH_SEPARATOR = "\\";

/** 
 * 指定したpathのファイルを指定したmodeで開き、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したファイルオブジェクトを返します。
 * @param {String} path ファイルのパスを文字列で指定します。
 * @param {String} mode ファイルを読み書きするためのモードを指定します。
 * @param {Function} block ブロック
 * <a href = "File.html#constructor">コンストラクタ</a>参照
 */
File.open = function(path, mode, block) {
  var file = new File(path, mode);
  if(block) {
    try {
      block(file);
    } finally {
      file.close();
    }
  } else {
    return file;
  }
};

/** 
 * fromファイルの名前をtoに変更します。
 * @param {String} from 変更前のファイル名
 * @param {String} to 変更後のファイル名
 * @throws ファイルが存在しない場合にスローされます。
 */
File.rename = function(from, to) {
  if(Const.FSO.FolderExists(from)) {
    Const.FSO.MoveFolder(from, to);
  } else if(Const.FSO.FileExists(from)) {
    Const.FSO.MoveFile(from, to);
  } else {
    throw new Error(-1, "File Not Found: " + from);
  }
};

/** 
 * fromファイルをtoにコピーします。
 * @param {String} from コピー元ファイル名
 * @param {String} to コピー先ファイル名
 * @throws ファイルが存在しない場合にスローされます。
 */
File.copy = function(from, to) {
  if(Const.FSO.FolderExists(from)) {
    Const.FSO.CopyFolder(from, to);
  } else if(Const.FSO.FileExists(from)) {
    Const.FSO.CopyFile(from, to);
  } else {
    throw new Error(-1, "File Not Found: " + from);
  }
};

/** 
 * ファイル、または、ディレクトリが存在するかどうか判定します。
 * @param {String} path ファイル、または、ディレクトリのパス
 * @return {Boolean} ファイル、または、ディレクトリが存在する場合はtrue、そうでない場合はfalseを返します。
 */
File.exist = function(path) {
  if(Const.FSO.FolderExists(path)) {
    return true;
  } else if(Const.FSO.FileExists(path)) {
    return true;
  } else {
    return false;
  }
};

/** 
 * 指定したpathのファイル、または、ディレクトリを削除します。
 * @param {String} path ファイル、または、ディレクトリのパス
 * @param {Boolean} force 読み取り専用も削除するかどうか
 * @throws ファイルが存在しない場合にスローされます。
 */
File.unlink = function(path, force) {
  if(!force) {
    force = false;
  }
  if(Const.FSO.FolderExists(path)) {
    Const.FSO.DeleteFolder(path, force);
  } else if(Const.FSO.FileExists(path)) {
    Const.FSO.DeleteFile(path, force);
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

/** 
 * 指定したpathのファイル、または、ディレクトリのサイズを取得します。
 * @param {String} path ファイル、または、ディレクトリのパス
 * @return {Number} ファイル、または、ディレクトリのサイズ(Byte単位)
 * @throws ファイルが存在しない場合にスローされます。
 */
File.size = function(path) {
  if(Const.FSO.FolderExists(path)) {
    return Const.FSO.GetFolder(path).Size;
  } else if(Const.FSO.FileExists(path)) {
    return Const.FSO.GetFile(path).Size;
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

/** 
 * 指定したbasedirからの相対パスで、pathnameのフルパスを取得します。
 * @param {String} pathname ファイル、または、ディレクトリのパス
 * @param {String} basedir (オプション)起点となるディレクトリ(初期値 カレントディレクトリ)
 * @return {String} ファイル、または、ディレクトリのフルパス
 */
File.realpath = function(pathname, basedir) {
  var myBasedir = basedir || Dir.getwd();

  Dir.chdir(myBasedir, function(path) {
    realpath = Const.FSO.GetAbsolutePathName(pathname);
  });

  return realpath;
};

/** 
 * 指定したfilenameの拡張子を取得します。
 * @param {String} filename ファイルのパス
 * @return {String} filenameの拡張子(ドット付き; .ext)
 * @example 使用例：
print(File.extname("c:\work")); // ""
print(File.extname("c:\work\build_y2.xml")); // ".xml"
 */
File.extname = function(filename) {
  var extensionName = Const.FSO.GetExtensionName(filename);
  if(extensionName === "") {
    return extensionName;
  } else {
    return "." + extensionName;
  }
};

/** 
 * 指定したfilenameのパスを除いたファイル名を取得します。
 * @param {String} filename ファイルのパス
 * @return {String} パスを除いたファイル名(/path/to/file.txt → file.txt)
 */
File.basename = function(filename) {
  var file = Const.FSO.GetFile(filename);
  return file.Name;
};

/** 
 * 一時ファイルを作成します。(ファイル名:MyDocuments/YYYYMMDDHH24MISSt)
 * @param {Function} block ブロック
 */
File.createTemporaryFile = function(block) {
  var tempFilePath = SpecialFolders.getMyDocuments() + "\\" + new Date().getYYYYMMDDHH24MISS();
  File.open(tempFilePath, "w", function(file) {
    block(file);
  });

  // ファイルが存在していれば削除する。
  if(File.exist(tempFilePath)) {
    File.unlink(tempFilePath);
  }
};

/**
 * 指定されたファイルが最後にアクセスされたときの日付と時刻を返します。
 * @param {String} filename ファイルのパス
 * @return {Date} 指定されたファイルが最後にアクセスされたときの日付と時刻
 * @throws ファイルが存在しない場合にスローされます。
 */
File.atime = function(filename) {
  if(Const.FSO.FileExists(filename)) {
    var file = Const.FSO.GetFile(filename);
    return file.DateLastAccessed;
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

/**
 * 指定されたファイルが最後に更新されたときの日付と時刻を返します。
 * @param {String} filename ファイルのパス
 * @return {Date} 指定されたファイルが最後に更新されたときの日付と時刻
 * @throws ファイルが存在しない場合にスローされます。
 */
File.mtime = function(filename) {
  if(Const.FSO.FileExists(filename)) {
    var file = Const.FSO.GetFile(filename);
    return file.DateLastModified;
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

/**
 * 指定されたファイルが作成されたときの日付と時刻を返します。
 * @param {String} filename ファイルのパス
 * @return {Date} 指定されたファイルが作成されたときの日付と時刻
 * @throws ファイルが存在しない場合にスローされます。
 */
File.ctime = function(filename) {
  if(Const.FSO.FileExists(filename)) {
    var file = Const.FSO.GetFile(filename);
    return file.DateCreated;
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

/**
 * filenameの一番後ろのスラッシュより前を文字列として返します。スラッシュを含まないファイル名に対しては"."(カレントディレクトリ)を返します。
 * @param {String} filename ファイルのパス
 * @return {String} 指定されたファイルのdirname
 */
File.dirname = function(filename) {
  if(filename === "/") {
    return "/";
  } else if(filename.indexOf("/") === -1) {
    return ".";
  } else {
    var targetString = filename;
    if(filename.endsWith("/")) {
      targetString = targetString.substring(0, targetString.length - 1);
    }

    while(true) {
      var slashIndex = targetString.lastIndexOf("/");
      if(slashIndex === 0) {
        return "/";
      } else {
        var result = targetString.substring(0, slashIndex);
        if(result.charAt(result.length - 1) !== '/') {
          return result;
        }
        targetString = result;
      }
    }
  }
};

/**
 * 指定されたファイルの8.3形式のファイル名を返します。
 * @param {String} filename ファイルのパス
 * @return {String} 指定されたファイルの8.3形式のファイル名
 * @throws ファイルが存在しない場合にスローされます。
 */
File.getShortName = function(filename) {
  if(Const.FSO.FileExists(filename)) {
    var file = Const.FSO.GetFile(filename);
    return file.ShortName;
  } else {
    throw new Error(-1, "File Not Found: " + path);
  }
};

// Prototypes of File
File.prototype = {
  /** 
   * ファイルから一行ずつ読み取り、ブロックを呼び出します。
   * @param {Function} block ブロック
   */
  each: function(block) {
    while(!this.ts.AtEndOfStream) {
      block(this.ts.ReadLine());
    }
  },
  /**
   * ファイルから一行読み取ります。
   * @return {String} 読み取った文字列
   */
  readLine: function() {
    return this.ts.ReadLine();
  },
  /** 
   * ファイルへ一行書き込みます。
   * @param {String} line 書きこむ文字列
   */
  puts: function(line) {
    this.ts.WriteLine(line);
  },
  /** 
   * ファイルへ書き込みます。
   * @param {String} line 書きこむ文字列
   */
  write: function(line) {
    this.ts.Write(line);
  },
  /** 
   * ファイルをクローズします。
   */
  close: function() {
    this.ts.Close();
  },
  /**
   * このファイルのパスを取得します。
   * @return {String} このファイルのパス
   */
  getPath: function() {
    return this.path;
  },
  /**
   * 指定されたファイルが最後にアクセスされたときの日付と時刻を返します。
   * @return {Date} 指定されたファイルが最後にアクセスされたときの日付と時刻
   */
  atime: function() {
    return File.atime(this.getPath());
  },
  /**
   * 指定されたファイルが最後に更新されたときの日付と時刻を返します。
   * @return {Date} 指定されたファイルが最後に更新されたときの日付と時刻
   */
  mtime: function() {
    return File.mtime(this.getPath());
  },
  /**
   * 指定されたファイルが作成されたときの日付と時刻を返します。
   * @return {Date} 指定されたファイルが作成されたときの日付と時刻
   */
  ctime: function() {
    return File.ctime(this.getPath());
  },
  /**
   * 指定されたファイルの8.3形式のファイル名を返します。
   * @return {Date} 指定されたファイルの8.3形式のファイル名
   */
  getShortName: function() {
    return File.getShortName(this.getPath());
  },
  /**
   * ファイルが終端に達しているかどうかを返します。
   * @return {Boolean} ファイルが終端に達している場合はtrue、それ以外の場合はfalseを返します。
   */
  eof: function() {
    return this.ts.AtEndOfStream;
  },
  /**
   * ファイルから全ての内容を読み取ります。
   * @return {String} 読み取った文字列
   */
  read: function() {
    return this.ts.ReadAll();
  }
};
/** 
 * pathの存在チェックを行った上で、ディレクトリオブジェクトを作成します。
 * @class ディレクトリの操作を行うためのクラスです。
<pre class = "code">
使用例：
// src以下のファイル一覧を表示
var d = new Dir("src");
d.each(function(item) {
  print(item);
});
</pre>
 * @param  {String} path ディレクトリのパスを文字列で指定します。
 * @throws ディレクトリが存在しない場合にスローされます。
 */
var Dir = function(path) {
  if(!Dir.exist(path)) {
    throw new Error(path + " not found or is not directory.");
  }
  this.path = path;
};

/**
 * カレントディレクトリのフルパスを文字列で返します。
 * @return {String} カレントディレクトリのフルパス
 */
Dir.getwd = function() {
  return Const.WSHELL.CurrentDirectory;
};

/**
 * <a href = "#.getwd">Dir#getwd</a>のエイリアスです。
 * @function
 */
Dir.pwd = Dir.getwd;

/**
 * カレントディレクトリをpathに変更します。
 * ブロックが指定された場合、カレントディレクトリの変更はブロックの実行中に限られます。
 * @param {String} path ディレクトリのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @example 使用例：
Dir.chdir("src", function(path) {
  print(File.size("date.js"));
});

print(File.size("mado.js"));
 */
Dir.chdir = function(path, block) {
  if(!path) {
    path = Const.INITIAL_CURRENT_DIRECTORY;
  } else if(!File.exist(path)) {
    throw new Error(-1, "Path Not Found: " + path);
  }

  if(block) {
    // 現在のワーキングディレクトリを取得
    var cwd = Dir.getwd();

    // 指定のパスにディレクトリ移動
    Const.WSHELL.CurrentDirectory = path;

    // ブロックを実行
    block(path);

    // ワーキングディレクトリを元に戻す
    Const.WSHELL.CurrentDirectory = cwd;
  } else {
    Const.WSHELL.CurrentDirectory = path;
  };
};

/**
 * pathで指定された新しいディレクトリを作ります。
 * @param {String} path ディレクトリのパスを文字列で指定します。
 * @return {Object} 作成したディレクトリのディレクトリオブジェクト
 */
Dir.mkdir = function(path) {
  Const.FSO.CreateFolder(path);
  return new Dir(path);
};

/**
 * pathで指定されたディレクトリの親ディレクトリのパスを文字列で取得します。
 * @param {String} path ディレクトリのパスを文字列で指定します。
 * @return {String} 親ディレクトリのパス
 */
Dir.getParentDir = function(path) {
  return Const.FSO.GetParentFolderName(path);
};

Dir.open = function(path, block) {
  if(block) {
    var dir = new Dir(path);
    block(dir);
  } else {
    return new Dir(path);
  }
};

/**
 * file_nameで与えられたディレクトリが存在するかどうかをチェックします。
 * @param {String} file_name ディレクトリのパスを文字列で指定します。
 * @return {Boolean} 存在する場合はtrue、存在しない場合はfalseを返します。
 */
Dir.exist = function(file_name) {
  return Const.FSO.FolderExists(file_name);
};

/**
 * ディレクトリpathに含まれるファイルエントリ名の配列を返します。
 * @param {String} path ディレクトリのパスを文字列で指定します。
 * @return {Boolean} ファイルエントリ名の配列
 */
Dir.entries = function(path) {
  var dir = new Dir(path);

  var result = [];

  dir.each(function(item) {
    result.push(item);
  });

  return result;
};

/**
 * ディレクトリpathに含まれる(サブディレクトリも含む)ファイルエントリ名の配列を返します。
 * patternを指定した場合、patternの正規表現にマッチするファイルのみ抽出します。
 * @param {String} path ディレクトリのパスを文字列で指定します。
 * @param {String} pattern 抽出するファイル名の正規表現パターン
 * @param {Function} block ブロック
 * @return {Boolean} ファイルエントリ名の配列
 * @example 使用例：
// docs以下のhtmlファイルのみ抽出
Dir.find("docs", ".*.html$", function(item) {
  print(item);
});
 */
Dir.find = function(path, pattern, block) {
  var dir = new Dir(path);

  var result = [];

  var processFile = function(item) {
    if(pattern === undefined || item.match(pattern)) {
      if(isFunction(block)) {
        block(item);
      } else {
        result.push(item);
      }
    }
  };

  dir.each(function(item) {
    if(Dir.exist(item)) {
      var subresult = Dir.find(item, pattern, block);
      var subresultLength = subresult.length;
      for(var i=0; i<subresultLength; i++) {
        processFile(subresult[i]);
      }
    } else {
      processFile(item);
    }
  });

  return result;
};

// Prototypes of Dir
Dir.prototype = {
  /**
   * このオブジェクトが持つpathで指定されたディレクトリの親ディレクトリのパスを文字列で取得します。
   * @return {String} 親ディレクトリのパス
   */
  getParentDir: function() {
    return Dir.getParentDir(this.path);
  },
  /**
   * このオブジェクトが持つpathに含まれるファイルエントリ名の配列を返します。
   * @return {Boolean} ファイルエントリ名の配列
   */
  each: function(block) {
    var dir = Const.FSO.GetFolder(this.path);

    var enumDir = new Enumerator(dir.SubFolders);
    for (;!enumDir.atEnd(); enumDir.moveNext()) {
       block(enumDir.item().Path);
    }

    var enumFile = new Enumerator(dir.Files);
    for (;!enumFile.atEnd(); enumFile.moveNext()) {
       block(enumFile.item().Path);
    }
  }
};
/** 
 * ドライブの存在チェックを行った上で、ドライブオブジェクトを作成します。
 * @class ドライブの操作を行うためのクラスです。
<pre class = "code">
使用例：
// Cドライブの空き容量を表示
var d = new Drive("c:");
print(d.getFreeSpace());
</pre>
 * @param  {String} path ディレクトリのパスを文字列で指定します。
 * @throws ディレクトリが存在しない場合にスローされます。
 */
var Drive = function(path) {
  if(!Drive.exist(path)) {
    throw new Error("Drive " + path + " not found.");
  }
  this.path = path;
  this.driveObj = Const.FSO.GetDrive(path);
};

/**
 * ドライブタイプの定数：不明(0)
 */
Drive.DRIVE_TYPE_UNKNOWN = 0;

/**
 * ドライブタイプの定数：リムーバブルディスク(1)
 */
Drive.DRIVE_TYPE_REMOVABLE_DISK = 1;

/**
 * ドライブタイプの定数：ハードディスク(2)
 */
Drive.DRIVE_TYPE_HARD_DISK = 2;

/**
 * ドライブタイプの定数：ネットワークドライブ(3)
 */
Drive.DRIVE_TYPE_NETWORK_DRIVE = 3;

/**
 * ドライブタイプの定数：CD-ROM(4)
 */
Drive.DRIVE_TYPE_CDROM = 4;

/**
 * ドライブタイプの定数：RAMディスク(5)
 */
Drive.DRIVE_TYPE_RAM_DISK = 5;

/**
 * ドライブの一覧を取得します。
 * @return {Array} Driveオブジェクトの配列
 */
Drive.getDrives = function() {
  var result = [];
  var e = new Enumerator(Const.FSO.Drives);
  for (; !e.atEnd(); e.moveNext())
  {
    result.push(new Drive(e.item().Path));
  }

  return result;
};

/**
 * pathで与えられたドライブが存在するかどうかをチェックします。
 * @param {String} path ドライブのパスを文字列で指定します。
 * @return {Boolean} 存在する場合はtrue、存在しない場合はfalseを返します。
 */
Drive.exist = function(path) {
  return Const.FSO.DriveExists(path);
};

// Prototypes of Drive
Drive.prototype = {
  /**
   * このドライブのAvailableSpaceを取得します。
   * @return {Number} このドライブのAvailableSpace
   */
  getAvailableSpace: function() {
    return this.driveObj.AvailableSpace;
  },
  /**
   * このドライブのDriveLetterを取得します。
   * @return {String} このドライブのDriveLetter
   */
  getDriveLetter: function() {
    return this.driveObj.DriveLetter;
  },
  /**
   * このドライブのDriveTypeを取得します。
   * @return {Number} このドライブのDriveType
   */
  getDriveType: function() {
    return this.driveObj.DriveType;
  },
  /**
   * このドライブのFreeSpaceを取得します。
   * @return {Number} このドライブのFreeSpace
   */
  getFreeSpace: function() {
    return this.driveObj.FreeSpace;
  },
  /**
   * このドライブのFileSystemを取得します。
   * @return {String} このドライブのFileSystem
   */
  getFileSystem: function() {
    return this.driveObj.FileSystem;
  },
  /**
   * このドライブが使用可能かどうかを取得します。
   * @return {String} このドライブが使用可能な場合true、そうでない場合はfalseを返します。
   */
  isReady: function() {
    return this.driveObj.IsReady;
  },
  /**
   * このドライブのSerialNumberを取得します。
   * @return {String} このドライブのSerialNumber
   */
  getSerialNumber: function() {
    return this.driveObj.IsReady;
  },
  /**
   * このドライブのShareNameを取得します。
   * @return {String} このドライブのShareName
   */
  getShareName: function() {
    return this.driveObj.ShareName;
  },
  /**
   * このドライブのTotalSizeを取得します。
   * @return {String} このドライブのTotalSize
   */
  getTotalSize: function() {
    return this.driveObj.TotalSize;
  },
  /**
   * このドライブのVolumeNameを取得します。
   * @return {String} このドライブのVolumeName
   */
  getVolumeName: function() {
    return this.driveObj.VolumeName;
  }
};
/** 
 * hostとportを指定してHTTPオブジェクトを作成します。
 * @class HTTPの操作を行うためのクラスです。
<pre class = "code">
使用例：
HTTP.start("www.yahoo.co.jp", 80, function(http) {
  print(http.get("/index.html"));
});
</pre>
 * @param {String} host HTTP操作の対象ホスト
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 */
var HTTP = function(host, port) {
  this.host = host;
  this.port = port || DEFAULT_PORT_NUMBER;
};

/**
 * 定数：デフォルトのポート番号(80)
 */
HTTP.DEFAULT_PORT_NUMBER = 80;
/**
 * 定数：レスポンス形式（テキスト）
 */
HTTP.RESPONSE_TYPE_TEXT = "TEXT";
/**
 * 定数：レスポンス形式（ボディ）
 */
HTTP.RESPONSE_TYPE_BODY = "BODY";
/**
 * 定数：レスポンス形式（XML）
 */
HTTP.RESPONSE_TYPE_XML = "XML";
/**
 * 定数：XHR実装クラス：Microsoft.XMLHTTP
 */
HTTP.XHR_IMPL_CLASS_Microsoft_XMLHTTP = "Microsoft.XMLHTTP";
/**
 * 定数：XHR実装クラス：MSXML2.XMLHTTP5.0
 */
HTTP.XHR_IMPL_CLASS_MSXML2_XMLHTTP5_0 = "MSXML2.XMLHTTP5.0";

/**
 * グローバル変数：XHR実装クラス(初期値 XHR_IMPL_CLASS_Microsoft_XMLHTTP
 */
HTTP.XHR_IMPL_CLASS = HTTP.XHR_IMPL_CLASS_Microsoft_XMLHTTP;

/** 
 * hostとportを指定してHTTPオブジェクトを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したHTTPオブジェクトを返します。
 * @param {String} host HTTP操作の対象ホスト
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @param {Function} block (オプション)ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したHTTPオブジェクト
 */
HTTP.start = function(host, port, block) {
  var http = new HTTP(host, port);
  if(block) {
    block(http);
  } else {
    return http;
  }
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行います。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
 * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
 * @return {String} レスポンス
 * @example 使用例：
print(HTTP.get("www.yahoo.co.jp", "/index.html", 80));
 */
HTTP.get = function(address, path, port, header, responseType) {
  var myPort = port || DEFAULT_PORT_NUMBER;
  var myHeader = header || {};
  var myResponseType = responseType || HTTP.RESPONSE_TYPE_TEXT;
  var http = new HTTP(address, myPort);
  return http.get(path, myHeader, myResponseType);
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行い、printします。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
 * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
 * @example 使用例：
HTTP.get_print("www.yahoo.co.jp", "/index.html", 80);
 */
HTTP.get_print = function(address, path, port, header, responseType) {
  print(HTTP.get(address, path, port, header, responseType));
};

// Prototypes of HTTP
HTTP.prototype = {
  /** 
   * 指定されたpath、headerでGETリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
   * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  get: function(path, header, responseType, block) {
    return this.request("GET", path, header, responseType, "", block);
  },
  /** 
   * 指定されたpath、headerでPOSTリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ(初期値 [])
   * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
   * @param {Object} body (オプション)HTTPボディ(初期値 undefined)
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  post: function(path, header, responseType, body, block) {
    return this.request("POST", path, header, responseType, body, block);
  },
  // private
  request: function(method, path, header, responseType, body, block) {
    // パラメータの調整
    var myHeader = header || {};

    // リクエストの準備
    var xhr = new ActiveXObject(HTTP.XHR_IMPL_CLASS);
    xhr.open(method, "http://" + this.host + ":" + this.port + path, false);

    for(var headerName in myHeader) {
      if(myHeader.hasOwnProperty(headerName)) {
        xhr.setRequestHeader(headerName, myHeader[headerName]);
      }
    }

    // リクエストの送信
    xhr.send(body);

    // レスポンスの処理
    if(xhr.readyState === 4) {
      if(xhr.status === 200) {
        var response;
        if(responseType === HTTP.RESPONSE_TYPE_TEXT) {
          response = xhr.responseText;
        } else if(responseType === HTTP.RESPONSE_TYPE_BODY) {
          response = xhr.responseBody;
        } else if(responseType === HTTP.RESPONSE_TYPE_XML) {
          response = xhr.responseXML;
        } else {
          response = xhr.responseText;
        }
        if(isFunction(block)) {
          block(response);
        } else {
          return response;
        }
      } else {
        throw new Error(xhr.statusText);
      }
    }
  }
};
/**
 * インスタンス化しません。
 * @class FTP通信に関する機能を提供するクラス
<pre class = "code">
使用例：
FTP.get("ftp.example.com", "user", "password", "/path/to/file", "test.log");
</pre>
 */
var FTP = {};


/**
 * FTP転送モードの定数：ASCIIモード(1)
 */
FTP.MODE_ASCII = 1;

/**
 * FTP転送モードの定数：BINARYモード(2)
 */
FTP.MODE_BINARY = 2;

/**
 * 指定したパラメータのファイルをFTPで取得します。
 * @param {String} host FTPサーバ
 * @param {String} userId FTPユーザID
 * @param {String} password FTPパスワード
 * @param {Number} mode FTP転送モード
 * @param {String} path GETするファイルのあるパス
 * @param {String} fileName GETするファイル名
 */
FTP.get = function(host, userId, password, mode, path, fileName) {
  File.createTemporaryFile(function(file) {
    file.puts("open " + host);
    file.puts(userId);
    file.puts(password);
    file.puts("cd " + path);

    if(mode === FTP.MODE_ASCII) {
      file.puts("ascii");
    } else {
      file.puts("binary");
    }

    file.puts("get " + fileName);
    file.puts("bye");

    Process.exec("ftp", ["-s:\"" + file.getPath() + "\""], Const.WINDOW_STYLE_MIN, true);
  });
};

/**
 * 指定したパラメータのファイルをFTPでPUTします。
 * @param {String} host FTPサーバ
 * @param {String} userId FTPユーザID
 * @param {String} password FTPパスワード
 * @param {Number} mode FTP転送モード
 * @param {String} path PUTする先のパス
 * @param {String} fileName PUTするファイル名
 */
FTP.put = function(host, userId, password, mode, path, fileName) {
  File.createTemporaryFile(function(file) {
    file.puts("open " + host);
    file.puts(userId);
    file.puts(password);
    file.puts("cd " + path);

    if(mode === FTP.MODE_ASCII) {
      file.puts("ascii");
    } else {
      file.puts("binary");
    }

    file.puts("put " + fileName);
    file.puts("bye");

    Process.exec("ftp", ["-s:\"" + file.getPath() + "\""], Const.WINDOW_STYLE_MIN, true);
  });
};
/** 
 * 新しいクリップボードを作成する
 * @class InternetExplorerを利用した、クリップボードへの格納、および、取得を行うクラス。
 */
var ClipboardIE = function() {
  this._ie = new ActiveXObject('InternetExplorer.Application');
  this._ie.navigate("about:blank");
  while(this._ie.Busy) {
    sleep(10);
  }
  this._ie.Visible = false;
  this._textarea = this._ie.document.createElement("textarea");
  this._ie.document.body.appendChild(this._textarea);
  this._textarea.focus();
  this._closed = false;
};

/** 
 * 新しいクリップボードを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したクリップボードを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
ClipboardIE.open = function(block) {
  if (!isFunction(block)) {
    return new ClipboardIE();
  }

  try {
    var clip = new ClipboardIE();
    block(clip);
  } finally {
    if(clip) {
      clip.close();
    }
  }
};

// Prototypes of ClipboardIE
ClipboardIE.prototype = {
  /**
   * クリップボードへ文字列を格納します。
   * @param {String} text 格納する文字列
   */
  set: function(text) {
    if (this._closed) {
      return;
    }
    this._textarea.innerText = text;
    this._ie.execWB(17 /* select all */, 0);
    this._ie.execWB(12 /* copy */, 0);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    if (this._closed) {
      return;
    }
    this._textarea.innerText = "";
    this._ie.execWB(13 /* paste */, 0);
    return this._textarea.innerText;
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this._ie.Quit();
    this._closed = true;
  }
};
/** 
 * 新しいクリップボードを作成する
 * @class Excelを利用した、クリップボードへの格納、および、取得を行うクラス。
 */
var ClipboardExcel = function() {
  this.excel = Excel.create();
  this.sheet = this.excel.getSheetByIndex(0);
  this.sheet.setFormat(1, 1, "@");
};

/** 
 * 新しいクリップボードを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したクリップボードを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
ClipboardExcel.open = function(block) {
  if (!isFunction(block)) {
    return new ClipboardExcel();
  }

  try {
    var clip = new ClipboardExcel();
    block(clip);
  } finally {
    if(clip) {
      clip.close();
    }
  }
};

// Prototypes of ClipboardExcel
ClipboardExcel.prototype = {
  /**
   * クリップボードへ文字列を格納します。
   * @param {String} text 格納する文字列
   */
  set: function(text) {
    this.sheet.setValue(1, 1, text);
    this.sheet.copy(1, 1);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    return this.sheet.getValue(1, 1);
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this.excel.quitDiscardChanges();
  }
};
/** 
 * 新しいクリップボードを作成する
 * @class クリップボードへの格納、および、取得を行うクラス。
<pre class = "code">
使用例：
// "日本語"という文字列をクリップボードへコピーして、ペーストします。
Clipboard.open(function(clip) {
  clip.set("日本語");
  sendKeys("^v");
});
</pre>
 */
var Clipboard = function() {
  if(Excel.available()) {
    this.clip = new ClipboardExcel();
  } else {
    this.clip = new ClipboardIE();
  }
};

/** 
 * 新しいクリップボードを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したクリップボードを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Clipboard.open = function(block) {
  if (!isFunction(block)) {
    return new Clipboard();
  }

  try {
    var clip = new Clipboard();
    block(clip);
  } finally {
    if(clip) {
      clip.close();
    }
  }
};

// Prototypes of Clipboard
Clipboard.prototype = {
  /**
   * クリップボードへ文字列を格納します。
   * @param {String} text 格納する文字列
   */
  set: function(text) {
    this.clip.set(text);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    return this.clip.get();
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this.clip.close();
  }
};
/** 
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったデータベースへの接続を行うクラス。<br/>
          通常はこのクラスを直接使用せず、サブクラスを使います。
 * @param {String} connectString 接続文字列
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoConnection = function(connectString, userName, password) {
  this.con = new ActiveXObject("ADODB.Connection");
  this.con.Open(connectString, userName, password);
};

/**
 * Schema Enumの定数：Table(20)
 */
AdoConnection.AD_SCHEMA_TABLES = 20;

/**
 * DataTypeEnumの定数：adArray(0x2000)
 */
AdoConnection.AD_DATATYPE_ARRAY = 0x2000;
/**
 * DataTypeEnumの定数：adBigInt(20)
 */
AdoConnection.AD_DATATYPE_BIGINT = 20;
/**
 * DataTypeEnumの定数：adBinary(128)
 */
AdoConnection.AD_DATATYPE_BINARY = 128;
/**
 * DataTypeEnumの定数：adBoolean(11)
 */
AdoConnection.AD_DATATYPE_BOOLEAN = 11;
/**
 * DataTypeEnumの定数：adBSTR(8)
 */
AdoConnection.AD_DATATYPE_BSTR = 8;
/**
 * DataTypeEnumの定数：adChapter(136)
 */
AdoConnection.AD_DATATYPE_CHAPTER = 136;
/**
 * DataTypeEnumの定数：adChar(129)
 */
AdoConnection.AD_DATATYPE_CHAR = 129;
/**
 * DataTypeEnumの定数：adCurrency(6)
 */
AdoConnection.AD_DATATYPE_CURRENCY = 6;
/**
 * DataTypeEnumの定数：adDate(7)
 */
AdoConnection.AD_DATATYPE_DATE = 7;
/**
 * DataTypeEnumの定数：adDBDate(133)
 */
AdoConnection.AD_DATATYPE_DBDATE = 133;
/**
 * DataTypeEnumの定数：adDBTime(134)
 */
AdoConnection.AD_DATATYPE_DBTIME = 134;
/**
 * DataTypeEnumの定数：adDBTimeStamp(135)
 */
AdoConnection.AD_DATATYPE_DBTIMESTAMP = 135;
/**
 * DataTypeEnumの定数：adDecimal(14)
 */
AdoConnection.AD_DATATYPE_DECIMAL = 14;
/**
 * DataTypeEnumの定数：adDouble(5)
 */
AdoConnection.AD_DATATYPE_DOUBLE = 5;
/**
 * DataTypeEnumの定数：adEmpty(0)
 */
AdoConnection.AD_DATATYPE_EMPTY = 0;
/**
 * DataTypeEnumの定数：adError(10)
 */
AdoConnection.AD_DATATYPE_ERROR = 10;
/**
 * DataTypeEnumの定数：adFileTime(64)
 */
AdoConnection.AD_DATATYPE_FILETIME = 64;
/**
 * DataTypeEnumの定数：adGUID(72)
 */
AdoConnection.AD_DATATYPE_GUID = 72;
/**
 * DataTypeEnumの定数：adIDispatch(9)
 */
AdoConnection.AD_DATATYPE_IDISPATCH = 9;
/**
 * DataTypeEnumの定数：adInteger(3)
 */
AdoConnection.AD_DATATYPE_INTEGER = 3;
/**
 * DataTypeEnumの定数：adIUnknown(13)
 */
AdoConnection.AD_DATATYPE_IUNKNOWN = 13;
/**
 * DataTypeEnumの定数：adLongVarBinary(205)
 */
AdoConnection.AD_DATATYPE_LONGVARBINARY = 205;
/**
 * DataTypeEnumの定数：adLongVarChar(201)
 */
AdoConnection.AD_DATATYPE_LONGVARCHAR = 201;
/**
 * DataTypeEnumの定数：adLongVarWChar(203)
 */
AdoConnection.AD_DATATYPE_LONGVARWCHAR = 203;
/**
 * DataTypeEnumの定数：adNumeric(131)
 */
AdoConnection.AD_DATATYPE_NUMERIC = 131;
/**
 * DataTypeEnumの定数：adPropVariant(138)
 */
AdoConnection.AD_DATATYPE_PROPVARIANT = 138;
/**
 * DataTypeEnumの定数：adSingle(4)
 */
AdoConnection.AD_DATATYPE_SINGLE = 4;
/**
 * DataTypeEnumの定数：adSmallInt(2)
 */
AdoConnection.AD_DATATYPE_SMALLINT = 2;
/**
 * DataTypeEnumの定数：adTinyInt(16)
 */
AdoConnection.AD_DATATYPE_TINYINT = 16;
/**
 * DataTypeEnumの定数：adUnsignedBigInt(21)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDBIGINT = 21;
/**
 * DataTypeEnumの定数：adUnsignedInt(19)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDINT = 19;
/**
 * DataTypeEnumの定数：adUnsignedSmallInt(18)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDSMALLINT = 18;
/**
 * DataTypeEnumの定数：adUnsignedTinyInt(17)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDTINYINT = 17;
/**
 * DataTypeEnumの定数：adUserDefined(132)
 */
AdoConnection.AD_DATATYPE_USERDEFINED = 132;
/**
 * DataTypeEnumの定数：adVarBinary(204)
 */
AdoConnection.AD_DATATYPE_VARBINARY = 204;
/**
 * DataTypeEnumの定数：adVarChar(200)
 */
AdoConnection.AD_DATATYPE_VARCHAR = 200;
/**
 * DataTypeEnumの定数：adVariant(12)
 */
AdoConnection.AD_DATATYPE_VARIANT = 12;
/**
 * DataTypeEnumの定数：adVarNumeric(139)
 */
AdoConnection.AD_DATATYPE_VARNUMERIC = 139;
/**
 * DataTypeEnumの定数：adVarWChar(202)
 */
AdoConnection.AD_DATATYPE_VARWCHAR = 202;
/**
 * DataTypeEnumの定数：adWChar(130)
 */
AdoConnection.AD_DATATYPE_WCHAR = 130;

/** 
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} connectString 接続文字列
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoConnection.open = function(connectString, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoConnection(connectString, userName, password);
  }

  try {
    var con = new AdoConnection(connectString, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoConnection
AdoConnection.prototype = {
  // private
  getRawConnection: function() {
    return this.con;
  },
  /**
   * コネクションの接続文字列を取得する。
   * @return {String} コネクションの接続文字列
   */
  getConnectionString: function() {
    return this.con.ConnectionString;
  },
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql) {
    try {
      // SQLの実行
      var rs = this.con.Execute(sql);
      return this.convertToArrayRs(rs);
    } catch(e) {
      throw e;
    } finally {
      if(rs) {
        rs.Close();
      }
    }
  },
  /**
   * 指定されたsqlを実行する。
   * @param {String} sql 実行するSQL
   */
  executeUpdate: function(sql) {
    // SQLの実行
    this.con.Execute(sql);
  },
  // private
  // Recordsetを配列にコンバートする。
  convertToArrayRs: function(rs) {
    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    while(!rs.EOF) {
      var record = {};
      for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
        var field = fe.item();
        record[field.Name] = field.Value;
      }
      result.push(record);

      rs.MoveNext();
    }

    return result;
  },
  /**
   * 指定されたsqlを実行するコマンドを作成します。
   * @param {String} sql 実行するSQL
   * @return {Object} コマンドオブジェクト
   */
  createCommand: function(sql) {
    return new AdoCommand(this, sql);
  },
  /**
   * データベースの接続を閉じる。
   */
  close: function() {
    this.con.Close();
  },
  /**
   * トランザクションを開始します。
   * @return {Number} トランザクションネストレベル
   */
  beginTrans: function() {
    return this.con.BeginTrans();
  },
  /**
   * トランザクションをコミットします。
   */
  commitTrans: function() {
    this.con.CommitTrans();
  },
  /**
   * トランザクションをロールバックします。
   */
  rollbackTrans: function() {
    this.con.RollbackTrans();
  },
  /**
   * テーブル名一覧を取得します。
   * @param {String} schema (オプション)スキーマを指定します。
   * @return {Array} テーブル名の配列を返します。
   */
  getTableNames: function(schema) {
    var rs = this.con.OpenSchema(AdoConnection.AD_SCHEMA_TABLES, arrayToSafeArray([undefined, schema, undefined, "TABLE"]));
    var array = this.convertToArrayRs(rs);
    var result = [];
    array.each(function(elem) {
      result.push(elem["TABLE_NAME"]);
    });

    return result;
  },
  /**
   * 指定したテーブルのフィールド名一覧を取得します。
   * @param {String} tableName フィールド名一覧を取得するテーブル名を指定します。
   * @return {Array} フィールド名一覧を返します。
   */
  getFieldNames: function(tableName) {
    // フィールドリスト取得のためのSQL空実行
    var rs = this.con.Execute("SELECT * FROM " + tableName);

    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
      result.push(fe.item().Name);
    }

    return result;
  },
  /**
   * 指定したテーブルのフィールド情報一覧を取得します。[{カラム名:カラムタイプ}]
   * @param {String} tableName フィールド情報一覧を取得するテーブル名を指定します。
   * @return {Array} フィールド情報一覧を返します。
   */
  getFieldInfo: function(tableName) {
    // フィールドリスト取得のためのSQL空実行
    var rs = this.con.Execute("SELECT * FROM " + tableName);

    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
      var item = fe.item();
      result.push({'Name': item.Name, 'Type': item.Type});
    }

    return result;
  }
};
/** 
 * 新しいADOのコマンドを作成する。<br/>
   通常、AdoCommandを直接インスタンス化せず、AdoConnectionのサブクラスから、createCommandを呼びます。
 * @class ADOコマンドを表すクラス。クエリにパラメータが必要な場合は、このクラスを使います。
<pre class = "code">
使用例：
AdoAccessConnection.open("c:\\path\\to\\test.mdb", "", "", function(con) {
  var command = con.createCommand("UPDATE TEST SET COL2 = ? WHERE COL1 = ?");
  command.setStringParameter("COL2", "えええ");
  command.setStringParameter("COL1", "Z");
  command.executeUpdate();
});

AdoAccessConnection.open("c:\\path\\to\\test.mdb", "", "", function(con) {
  var command = con.createCommand("SELECT COL2 FROM TEST WHERE COL1 = ?");
  command.setStringParameter("COL1", "Z");
  var rs = command.executeQuery();

  print(rs[0]["COL2"]);
});

</pre>
 * @param {Object} con Adoコネクション
 */
var AdoCommand = function(con, sql) {
  this.con = con;
  this.sql = sql;
  this.params = [];

  // コマンドの生成
  this.command = new ActiveXObject("ADODB.Command");
  this.command.CommandText = this.sql;
  this.command.CommandType = AdoCommand.COMMAND_TYPE_CMD_TEXT;
  this.command.ActiveConnection = this.con.getRawConnection();
};

/**
 * Command Typeの定数：テキスト(1)
 */
AdoCommand.COMMAND_TYPE_CMD_TEXT = 1; // 通常

/**
 * データ型の定数: Boolean(11)
 */
AdoCommand.DATA_TYPE_BOOLEAN = 11;

/**
 * データ型の定数: Char(129)
 */
AdoCommand.DATA_TYPE_CHAR = 129;

/**
 * データ型の定数: DBDate(133)
 */
AdoCommand.DATA_TYPE_DBDATE = 133;

/**
 * データ型の定数: DBTime(134)
 */
AdoCommand.DATA_TYPE_DBTIME = 134;

/**
 * データ型の定数: DBTimestamp(135)
 */
AdoCommand.DATA_TYPE_DBTIMESTAMP = 135;

/**
 * データ型の定数: Decimal(14)
 */
AdoCommand.DATA_TYPE_DECIMAL = 14;

/**
 * データ型の定数: Double(5)
 */
AdoCommand.DATA_TYPE_DOUBLE = 5;

/**
 * データ型の定数: Integer(3)
 */
AdoCommand.DATA_TYPE_INTEGER = 3;

/**
 * データ型の定数: Numeric(131)
 */
AdoCommand.DATA_TYPE_NUMERIC = 131;

/**
 * データ型の定数: Single(4)
 */
AdoCommand.DATA_TYPE_SINGLE = 4;

/**
 * パラメータ方向の定数: Input(1)
 */
AdoCommand.PARAMETER_DIRECTION_INPUT = 1;

// Prototypes of AdoCommand
AdoCommand.prototype = {
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @param {Array} params パラメータ
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql, params) {
    try {
      // パラメータの追加
      this.prepareParameter();

      // SQLの実行
      var rs = this.command.Execute();
      return this.con.convertToArrayRs(rs);
    } catch(e) {
      throw e;
    } finally {
      if(rs) {
        rs.Close();
      }
    }
  },
  /**
   * 指定されたsqlを実行する。
   * @param {String} sql 実行するSQL
   * @param {Array} params パラメータ
   */
  executeUpdate: function(sql, params) {
    // パラメータの追加
    this.prepareParameter();

    // SQLの実行
    this.command.Execute();
  },
  // private
  // パラメータの準備をする。
  prepareParameter: function() {
    // パラメータの追加
    var command = this.command;
    this.params.each(function(param) {
      var objParam = command.CreateParameter(
        param.Name
        , param.Type
        , param.Direction
        , param.Size
        , param.Value
      );
      command.Parameters.Append(objParam);
    });
  },
  /**
   * 論理型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Boolean} paramValue パラメータ値
   */
  setBooleanParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_BOOLEAN
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 文字列型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {String} paramValue パラメータ値
   */
  setStringParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_CHAR
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 日付型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setDateParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBDATE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 時刻型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setTimeParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBTIME
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Timestamp型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setTimestampParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBTIMESTAMP
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Deciaml型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Number} paramValue パラメータ値
   */
  setDecimalParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DECIMAL
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Double型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setDoubleParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DOUBLE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Integer型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setIntegerParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_INTEGER
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Numeric型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setNumericParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_NUMERIC
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Single型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setSingleParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_SINGLE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  }
};
/** 
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったAccessデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoAccessConnection.open("c:\\path\\to\\test.mdb", "hoge", "fuga", function(con) {
  var result = con.executeQuery("SELECT * FROM TEST");

  result.each(function(row) {
    print(row["COL1"]); // 'X'
  });
});

</pre>
 * @param {String} mdbFilePath mdbファイルのパス
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoAccessConnection = function(mdbFilePath, userName, password) {
  var connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbFilePath;
  this.con = new AdoConnection(connectString, userName, password).getRawConnection();
};

/** 
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} mdbFilePath mdbファイルのパス
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoAccessConnection.open = function(mdbFilePath, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoAccessConnection(mdbFilePath, userName, password);
  }

  try {
    var con = new AdoAccessConnection(mdbFilePath, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoAccessConnection
AdoAccessConnection.prototype = Object.create(AdoConnection.prototype);
AdoAccessConnection.prototype.constructor = AdoAccessConnection;
/**
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったOracleデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoOracleConnection.open("ORCL", "scott", "tiger", function(oracon) {
  var result = oracon.executeQuery("SELECT * FROM DUAL");

  result.each(function(row) {
    print(row["DUMMY"]); // 'X'
  });
});

</pre>
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoOracleConnection = function(dsName, userName, password) {
  var connectString = "Provider=MSDAORA;Data Source=" + dsName + ";";
  this.con = new AdoConnection(connectString, userName, password).getRawConnection();
};

/**
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoOracleConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOracleConnection(dsName, userName, password);
  }

  try {
    var con = new AdoOracleConnection(dsName, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOracleConnection.prototype = Object.create(AdoConnection.prototype);
AdoOracleConnection.prototype.constructor = AdoOracleConnection;
/**
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったODBCデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoOdbcConnection.open("DSN1", "user", "password", function(odbccon) {
  var result = odbccon.executeQuery("SELECT * FROM TEST");

  result.each(function(row) {
    print(row["COL1"]); // 'X'
  });
});

</pre>
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoOdbcConnection = function(dsName, userName, password) {
  var connectString = "Provider=MSDASQL;DSN=" + dsName + ";";
  this.con = new AdoConnection(connectString, userName, password).getRawConnection();
};

/**
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoOdbcConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOdbcConnection(dsName, userName, password);
  }

  try {
    var con = new AdoOdbcConnection(dsName, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOdbcConnection.prototype = Object.create(AdoConnection.prototype);
AdoOdbcConnection.prototype.constructor = AdoOdbcConnection;
/** 
 * 新しいキー送信クラスを作成する。
 * @class 連続で自動的にキー送信を行う
<pre class = "code">
使用例：
// ノートパッドを起動
Process.exec("notepad", [], Const.WINDOW_STYLE_NORMAL, false);

// 起動を待つ
sleep(100);

// "日本語"という文字列を書きこんで、abc.txtに保存。
var ks = new KeySender("無題")
    .sendJapaneseKeys("日本語")
    .sendKeyWithControl("s")
    .sendKeys("abc.txt")
    .sendEnter();
</pre>
 * @param {String} targetWindow KeySend対象のウインドウ名を指定します。
 * @param {Number} duration (オプション)KeySendごとの待ち時間(ミリ秒)を指定します(初期値 0)。
 */
var KeySender = function(targetWindow, duration) {
  this.myDuration = duration || 0;
  Const.WSHELL.AppActivate(targetWindow);
  sleep(KeySender.DURATION);
};

// Constants of KeySender
// private
// AppActivateしてからKeySendするまでの待ち時間
KeySender.DURATION = 10;

// Prototypes of KeySender
KeySender.prototype = {
  /**
   * 指定したtextをキー送信します。
   * @param {String} text 送信する文字列
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeys: function(text, num) {
    sendKeys(text, num, this.myDuration);
    return this;
  },
  /**
   * 指定した日本語textをキー送信します。
   * @param {String} text 送信する文字列
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendJapaneseKeys: function(text, num) {
    var myNum = num || 1;
    Clipboard.open(function(clip) {
      for(var i=0; i<myNum; i++) {
        clip.set(text);
        sendKeys("^v", 1, this.myDuration);
      }
    });
    return this;
  },
  /**
   * Shiftを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithShift: function(ch, num) {
    sendKeys("+" + ch, num, this.myDuration);
    return this;
  },
  /**
   * Controlを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithControl: function(ch, num) {
    sendKeys("^" + ch, num, this.myDuration);
    return this;
  },
  /**
   * Altを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithAlt: function(ch, num) {
    sendKeys("%" + ch, num, this.myDuration);
    return this;
  },
  /**
   * Back Spaceをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBackspace: function(ch, num) {
    sendKeys("{BACKSPACE}", num, this.myDuration);
    return this;
  },
  /**
   * Breakをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBreak: function(ch, num) {
    sendKeys("{BREAK}", num, this.myDuration);
    return this;
  },
  /**
   * Caps Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendCapsLock: function(ch, num) {
    sendKeys("{CAPSLOCK}", num, this.myDuration);
    return this;
  },
  /**
   * Deleteをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDelete: function(ch, num) {
    sendKeys("{DELETE}", num, this.myDuration);
    return this;
  },
  /**
   * ↓をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDownArrow: function(ch, num) {
    sendKeys("{DOWN}", num, this.myDuration);
    return this;
  },
  /**
   * Endをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnd: function(ch, num) {
    sendKeys("{END}", num, this.myDuration);
    return this;
  },
  /**
   * Enterをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnter: function(num) {
    sendKeys("{ENTER}", num, this.myDuration);
    return this;
  },
  /**
   * Escをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEsc: function(num) {
    sendKeys("{ESC}", num, this.myDuration);
    return this;
  },
  /**
   * Helpをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHelp: function(num) {
    sendKeys("{HELP}", num, this.myDuration);
    return this;
  },
  /**
   * Homeをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHome: function(num) {
    sendKeys("{HOME}", num, this.myDuration);
    return this;
  },
  /**
   * Insertをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendInsert: function(num) {
    sendKeys("{INSERT}", num, this.myDuration);
    return this;
  },
  /**
   * ←をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendLeftArrow: function(num) {
    sendKeys("{LEFT}", num, this.myDuration);
    return this;
  },
  /**
   * NumLockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendNumLock: function(num) {
    sendKeys("{NUMLOCK}", num, this.myDuration);
    return this;
  },
  /**
   * Page Downをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageDown: function(num) {
    sendKeys("{PGDN}", num, this.myDuration);
    return this;
  },
  /**
   * Page Upをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageUp: function(num) {
    sendKeys("{PGUP}", num, this.myDuration);
    return this;
  },
  /**
   * Print Screenをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPrintScreen: function(num) {
    sendKeys("{PRTSC}", num, this.myDuration);
    return this;
  },
  /**
   * →をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendRightArrow: function(num) {
    sendKeys("{RIGHT}", num, this.myDuration);
    return this;
  },
  /**
   * Scroll Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendScrollLock: function(num) {
    sendKeys("{SCROLLLOCK}", num, this.myDuration);
    return this;
  },
  /**
   * Tabをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendTab: function(num) {
    sendKeys("{TAB}", num, this.myDuration);
    return this;
  },
  /**
   * ↑をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendUpArrow: function(num) {
    sendKeys("{UP}", num, this.myDuration);
    return this;
  },
  /**
   * F1をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF1: function(num) {
    sendKeys("{F1}", num, this.myDuration);
    return this;
  },
  /**
   * F2をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF2: function(num) {
    sendKeys("{F2}", num, this.myDuration);
    return this;
  },
  /**
   * F3をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF3: function(num) {
    sendKeys("{F3}", num, this.myDuration);
    return this;
  },
  /**
   * F4をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF4: function(num) {
    sendKeys("{F4}", num, this.myDuration);
    return this;
  },
  /**
   * F5をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF5: function(num) {
    sendKeys("{F5}", num, this.myDuration);
    return this;
  },
  /**
   * F6をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF6: function(num) {
    sendKeys("{F6}", num, this.myDuration);
    return this;
  },
  /**
   * F7をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF7: function(num) {
    sendKeys("{F7}", num, this.myDuration);
    return this;
  },
  /**
   * F8をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF8: function(num) {
    sendKeys("{F8}", num, this.myDuration);
    return this;
  },
  /**
   * F9をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF9: function(num) {
    sendKeys("{F9}", num, this.myDuration);
    return this;
  },
  /**
   * F10をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF10: function(num) {
    sendKeys("{F10}", num, this.myDuration);
    return this;
  },
  /**
   * F11をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF11: function(num) {
    sendKeys("{F11}", num, this.myDuration);
    return this;
  },
  /**
   * F12をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF12: function(num) {
    sendKeys("{F12}", num, this.myDuration);
    return this;
  },
  /**
   * F13をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF13: function(num) {
    sendKeys("{F13}", num, this.myDuration);
    return this;
  },
  /**
   * F14をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF14: function(num) {
    sendKeys("{F14}", num, this.myDuration);
    return this;
  },
  /**
   * F15をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF15: function(num) {
    sendKeys("{F15}", num, this.myDuration);
    return this;
  },
  /**
   * F16をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF16: function(num) {
    sendKeys("{F16}", num, this.myDuration);
    return this;
  }
};
/**
 * インスタンス化しません。
 * @class プログラムの実行に関する機能を提供するクラス
<pre class = "code">
使用例：
// ノートパッドを起動
Process.exec("notepad", ["mado-debug.js"], Const.WINDOW_STYLE_NORMAL, false);
</pre>
 */
var Process = {};

/**
 * 指定したコマンドを実行します。
 * @param {String} command 実行するコマンド
 * @param {Array} args (オプション)引数(初期値 [])
 * @param {Number} windowStyle (オプション)<a href = "Const.html#.WINDOW_STYLE_MAX">ウインドウスタイル</a>参照(初期値 Const.WINDOW_STYLE_NORMAL)
 * @param {Boolean} waitOnReturn (オプション)trueの場合コマンドの実行が終わるまで待ちます。falseの場合は待ちません。(初期値 false)
 */
Process.exec = function(command, args, windowStyle, waitOnReturn) {
  var myArgs = args || [];
  var myWindowStyle = windowStyle || Const.WINDOW_STYLE_NORMAL;
  var myWaitOnReturn = waitOnReturn || false;

  var commandLine = command + " " + myArgs.join(" ");
  Const.WSHELL.Run(commandLine, myWindowStyle, myWaitOnReturn);
};
/**
 * インスタンス化しません。
 * @class レジストリにアクセスする機能を提供するクラス
<pre class = "code">
使用例：
// IP Messengerのユーザ名設定を取得
print(Registry.regRead("HKCU\\Software\\HSTools\\IPMsg\\NickNameStr"));

// "NickNameStra"に"hoge"を設定
Registry.regWrite("HKCU\\Software\\HSTools\\IPMsg\\NickNameStra", "hoge", Registry.REG_TYPE_REG_SZ);

// "NickNameStra"を削除
Registry.regDelete("HKCU\\Software\\HSTools\\IPMsg\\NickNameStra");
</pre>
 */
var Registry = {};

/**
 * レジストリのデータ型：REG_SZ
 */
Registry.REG_TYPE_REG_SZ = "REG_SZ";

/**
 * レジストリのデータ型：REG_DWORD
 */
Registry.REG_TYPE_REG_DWORD = "REG_DWORD";

/**
 * レジストリのデータ型：REG_BINARY
 */
Registry.REG_TYPE_REG_BINARY = "REG_BINARY";

/**
 * レジストリのデータ型：REG_EXPAND_SZ
 */
Registry.REG_TYPE_REG_EXPAND_SZ = "REG_EXPAND_SZ";

/**
 * レジストリ内のキー名または値名の値を返します。
 * @param {String} name 読み取るキーまたは値の名前です。
 * @return {Object} レジストリの値
 */
Registry.regRead = function(name) {
  return Const.WSHELL.RegRead(name);
};

/**
 * 新しいキーの作成、新しい値名の既存キーへの追加 (および値の設定)、既存の値名の値変更などを行います。
 * @param {String} name 作成、追加、変更するキー名、値名、または値を示す文字列値です。
 * @param {String} value 作成するキーの名前、既存のキーに追加する値の名前、または既存の値名に設定する値です。
 * @param {String} type (オプション)レジストリに保存する値のデータ型を指定します。
 */
Registry.regWrite = function(name, value, type) {
  var regType = type || Registry.REG_TYPE_REG_SZ;
  Const.WSHELL.RegWrite(name, value, regType);
};

/**
 * レジストリから指定されたキーまたは値を削除します。
 * @param {String} name レジストリ内で削除するキーまたは値の名前を示す文字列値です。
 */
Registry.regDelete = function(name) {
  Const.WSHELL.RegDelete(name);
};
/**
 * インスタンス化しません。
 * @class Excelの定数を定義するクラス。
 */
var ExcelColorConstants = {};

/**
 * カラーインデックス：1<span style = 'color: #000000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_01 = 1;

/**
 * カラーインデックス：2<span style = 'color: #FFFFFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_02 = 2;

/**
 * カラーインデックス：3<span style = 'color: #FF0000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_03 = 3;

/**
 * カラーインデックス：4<span style = 'color: #00FF00'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_04 = 4;

/**
 * カラーインデックス：5<span style = 'color: #0000FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_05 = 5;

/**
 * カラーインデックス：6<span style = 'color: #FFFF00'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_06 = 6;

/**
 * カラーインデックス：7<span style = 'color: #FF00FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_07 = 7;

/**
 * カラーインデックス：8<span style = 'color: #00FFFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_08 = 8;

/**
 * カラーインデックス：9<span style = 'color: #800000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_09 = 9;

/**
 * カラーインデックス：10<span style = 'color: #008000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_10 = 10;

/**
 * カラーインデックス：11<span style = 'color: #000080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_11 = 11;

/**
 * カラーインデックス：12<span style = 'color: #808000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_12 = 12;

/**
 * カラーインデックス：13<span style = 'color: #800080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_13 = 13;

/**
 * カラーインデックス：14<span style = 'color: #008080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_14 = 14;

/**
 * カラーインデックス：15<span style = 'color: #C0C0C0'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_15 = 15;

/**
 * カラーインデックス：16<span style = 'color: #808080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_16 = 16;

/**
 * カラーインデックス：17<span style = 'color: #9999FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_17 = 17;

/**
 * カラーインデックス：18<span style = 'color: #993366'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_18 = 18;

/**
 * カラーインデックス：19<span style = 'color: #FFFFCC'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_19 = 19;

/**
 * カラーインデックス：20<span style = 'color: #CCFFFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_20 = 20;

/**
 * カラーインデックス：21<span style = 'color: #660066'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_21 = 21;

/**
 * カラーインデックス：22<span style = 'color: #FF8080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_22 = 22;

/**
 * カラーインデックス：23<span style = 'color: #0066CC'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_23 = 23;

/**
 * カラーインデックス：24<span style = 'color: #CCCCFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_24 = 24;

/**
 * カラーインデックス：25<span style = 'color: #000080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_25 = 25;

/**
 * カラーインデックス：26<span style = 'color: #FF00FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_26 = 26;

/**
 * カラーインデックス：27<span style = 'color: #FFFF00'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_27 = 27;

/**
 * カラーインデックス：28<span style = 'color: #00FFFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_28 = 28;

/**
 * カラーインデックス：29<span style = 'color: #800080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_29 = 29;

/**
 * カラーインデックス：30<span style = 'color: #800000'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_30 = 30;

/**
 * カラーインデックス：31<span style = 'color: #008080'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_31 = 31;

/**
 * カラーインデックス：32<span style = 'color: #0000FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_32 = 32;

/**
 * カラーインデックス：33<span style = 'color: #00CCFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_33 = 33;

/**
 * カラーインデックス：34<span style = 'color: #CCFFFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_34 = 34;

/**
 * カラーインデックス：35<span style = 'color: #CCFFCC'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_35 = 35;

/**
 * カラーインデックス：36<span style = 'color: #FFFF99'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_36 = 36;

/**
 * カラーインデックス：37<span style = 'color: #99CCFF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_37 = 37;

/**
 * カラーインデックス：38<span style = 'color: #FF99CC'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_38 = 38;

/**
 * カラーインデックス：39<span style = 'color: #CC99FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_39 = 39;

/**
 * カラーインデックス：40<span style = 'color: #FFCC99'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_40 = 40;

/**
 * カラーインデックス：41<span style = 'color: #3366FF'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_41 = 41;

/**
 * カラーインデックス：42<span style = 'color: #33CCCC'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_42 = 42;

/**
 * カラーインデックス：43<span style = 'color: #99CC00'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_43 = 43;

/**
 * カラーインデックス：44<span style = 'color: #FFCC00'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_44 = 44;

/**
 * カラーインデックス：45<span style = 'color: #FF9900'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_45 = 45;

/**
 * カラーインデックス：46<span style = 'color: #FF6600'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_46 = 46;

/**
 * カラーインデックス：47<span style = 'color: #666699'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_47 = 47;

/**
 * カラーインデックス：48<span style = 'color: #969696'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_48 = 48;

/**
 * カラーインデックス：49<span style = 'color: #003366'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_49 = 49;

/**
 * カラーインデックス：50<span style = 'color: #339966'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_50 = 50;

/**
 * カラーインデックス：51<span style = 'color: #003300'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_51 = 51;

/**
 * カラーインデックス：52<span style = 'color: #333300'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_52 = 52;

/**
 * カラーインデックス：53<span style = 'color: #993300'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_53 = 53;

/**
 * カラーインデックス：54<span style = 'color: #993366'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_54 = 54;

/**
 * カラーインデックス：55<span style = 'color: #333399'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_55 = 55;

/**
 * カラーインデックス：56<span style = 'color: #333333'>■■■■■</span>
 */
ExcelColorConstants.COLOR_INDEX_56 = 56;
/** 
 * Excelオブジェクトを作成します。
 * @class Excelファイルの操作と読み書きを行うためのクラスです。<br/>
 * Excelがインストールされている必要があります。
<pre class = "code">
使用例：
// abc.xlsを開いて、Sheet1の、1行1列目の値を'hoge'に書き換えつつ、上書き保存する。
Excel.open("abc.xls", function(excel) {
  var sheet = excel.getSheetByName("Sheet1");
  sheet.setValue(1, 1, 'hoge');
  excel.save();
});
</pre>
 * @param {String} path Excelファイルのパスを文字列で指定します。指定しない場合は新しくブックを作成します。
 * @param {Boolean} fileType (オプション)Excel以外のファイルを開くときに、指定します。
 * @param {Boolean} columnInfo (オプション)Excel以外のファイルを開くときに、カラムの情報を指定します。
 * @throws pathがundefinedでない、かつ、ファイルが存在しない場合にスローされます。
 */
var Excel = function(path, fileType, columnInfo) {
  this.path = path;
  this.excelObj = new ActiveXObject("Excel.Application");

  if(path === undefined) {
    this.workbookObj = this.excelObj.Workbooks.Add();
  } else {
    if(!File.exist(path)) {
      throw new Error(-1, "File Not Found: " + path);
    }

    switch(fileType) {
      case Excel.IN_FILETYPE_CSV:
      case Excel.IN_FILETYPE_TSV:
        var isFileTypeCsv = fileType === Excel.IN_FILETYPE_CSV;
        var isFileTypeTsv = fileType === Excel.IN_FILETYPE_TSV;

        var columnInfoSafeArray = Excel.array2dToSafeArray2d(columnInfo);
        this.excelObj.Workbooks.OpenText(path, 1, 1, 1, 1, false, isFileTypeTsv, false, isFileTypeCsv, false, false, false, columnInfoSafeArray);
        this.workbookObj = this.excelObj.ActiveWorkbook;
        break;
      default:
        this.workbookObj = this.excelObj.Workbooks.Open(path);
    }
  }
};

/**
 * ファイルフォーマット：xlCSV(6)
 */
Excel.FILE_FORMAT_CSV = 6;

/**
 * ファイルフォーマット：xlCurrentPlatformText(-4158)
 */
Excel.FILE_FORMAT_TSV = -4158;

/**
 * ファイルフォーマット：xlWorkbookNormal(-4143)
 */
Excel.FILE_FORMAT_EXCEL = -4143;

/**
 * 入力ファイルタイプ：タブ区切り(1)
 */
Excel.IN_FILETYPE_TSV = 1;

/**
 * 入力ファイルタイプ：セミコロン(2)
 */
Excel.IN_FILETYPE_SEMI_COLON = 2;

/**
 * 入力ファイルタイプ：CSV(4)
 */
Excel.IN_FILETYPE_CSV = 4;

/**
 * 入力ファイルタイプ：スペース(8)
 */
Excel.IN_FILETYPE_SPACE = 8;

/**
 * 入力ファイルタイプ：その他(16)
 */
Excel.IN_FILETYPE_OTHER = 16;

/**
 * カラムタイプ：xlGeneralFormat(1)
 */
Excel.COLUMN_TYPE_GENERAL = 1;

/**
 * カラムタイプ：xlTextFormat(2)
 */
Excel.COLUMN_TYPE_TEXT = 2;

/**
 * カラムタイプ：xlMDYFormat(3)
 */
Excel.COLUMN_TYPE_MDY = 3;

/**
 * カラムタイプ：xlDMYFormat(4)
 */
Excel.COLUMN_TYPE_DMY = 4;

/**
 * カラムタイプ：xlYMDFormat(5)
 */
Excel.COLUMN_TYPE_YMD = 5;

/**
 * カラムタイプ：xlMYDFormat(6)
 */
Excel.COLUMN_TYPE_MYD = 6;

/**
 * カラムタイプ：xlDYMFormat(7)
 */
Excel.COLUMN_TYPE_DYM = 7;

/**
 * カラムタイプ：xlYDMFormat(8)
 */
Excel.COLUMN_TYPE_YDM = 8;

/**
 * カラムタイプ：xlSkipColumn(9)
 */
Excel.COLUMN_TYPE_SKIP = 9;

/**
 * カラムタイプ：xlEMDFormat(10)
 */
Excel.COLUMN_TYPE_EMD = 10;

/** 
 * Excelファイルを開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.open = function(path, block) {
  if(!isFunction(block)) {
    return new Excel(path);
  }

  try {
    var excel = new Excel(path);
    block(excel);
  } finally {
    if(excel) {
      excel.quit();
    }
  }
};

/** 
 * Excelファイルを読み取り専用(最後に変更を破棄)で開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.openReadonly = function(path, block) {
  if(!isFunction(block)) {
    return new Excel(path);
  }

  try {
    var excel = new Excel(path);
    block(excel);
  } finally {
    if(excel) {
      excel.quitDiscardChanges();
    }
  }
};

/** 
 * Excelファイルを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.create = function(block) {
  if(!isFunction(block)) {
    return new Excel();
  }

  try {
    var excel = new Excel();
    block(excel);
  } finally {
    if(excel) {
      excel.quit();
    }
  }
};

/** 
 * Excelファイルを読み取り専用(最後に変更を破棄)で作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.createReadonly = function(block) {
  if(!isFunction(block)) {
    return new Excel();
  }

  try {
    var excel = new Excel();
    block(excel);
  } finally {
    if(excel) {
      excel.quitDiscardChanges();
    }
  }
};

/**
 * プログラムを動作させる端末でExcelが利用可能かどうか。
 * @return {Boolean} Excelが利用可能な場合はtrue、それ以外の場合はfalseを返す。
 */
Excel.available = function() {
  try {
    var x = new ActiveXObject("Excel.Application");
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * ExcelブックをCSVファイルに変換する。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertToCsv = function(path) {
  Excel.openReadonly(path, function(excel) {
    excel.each(function(sheet) {
      sheet.activate();
      var outFileName = path + "_" + sheet.getName() + ".csv";
      excel.saveAsCsv(outFileName);
    });
  });
}

/**
 * Excelブックをタブ区切りファイルに変換する。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertToTsv = function(path) {
  Excel.openReadonly(path, function(excel) {
    excel.each(function(sheet) {
      sheet.activate();
      var outFileName = path + "_" + sheet.getName() + ".tsv";
      excel.saveAsTsv(outFileName);
    });
  });
}

/**
 * JScriptの二次元配列を、二次元のSafeArrayに変換する。
 * @param {Array} jsArray2d JScriptの二次元配列を指定します。
 * @return {Object} 変換された二次元のSafeArrayを返します。通常、VBArrayでラップして使います。<br/>
 *                  変換に失敗した場合はundefinedを返します。
 */
Excel.array2dToSafeArray2d = function(jsArray2d) {
  var safeArray = undefined;

  Excel.createReadonly(function(excel) {
    // 1シート目を取得
    var sheet = excel.getSheetByIndex(0);

    // 各セルに値を設定
    var i, j;
    for(i=0; i<jsArray2d.length; i++) {
      for(j=0; j<jsArray2d[i].length; j++) {
        var value = jsArray2d[i][j];
        sheet.setValue(i+1, j+1, value);
      }
    }

    // SafeArrayを取り出す。
    safeArray = sheet.getRawObject().Cells(1, 1).CurrentRegion.Value;
  });

  return safeArray;
}

/**
 * CSVファイルをExcelファイルに変換する。
 * @param {String} path CSVファイルのパスを文字列で指定します。
 * @param {Array} columnInfo (オプション)CSVファイルのカラム情報を指定します。<br>
<a href = "Excel.html#.COLUMN_TYPE_GENERAL">カラムタイプ</a>参照(初期値 全列に対しConst.COLUMN_TYPE_TEXT)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 * @example 使用例：
// 3カラム全て文字列として取り込む
var columnInfo = [
	[1, Excel.COLUMN_TYPE_TEXT]
	, [2, Excel.COLUMN_TYPE_TEXT]
	, [3, Excel.COLUMN_TYPE_TEXT]
];

// 拡張子がCSVだと、Excelが自動的に変換してうまくいかない
Excel.convertFromCsv("csv.txt", columnInfo);
 */
Excel.convertFromCsv = function(path, columnInfo) {
  var myPath = path;

  // 拡張子が.csvの場合は、うまく変換できないので、拡張子に.txtを付与したファイルにコピーする。
  if(path.endsWith(".csv")) {
    myPath = path + ".txt";
    File.copy(path, myPath);
  }

  var myColumnInfo = columnInfo;
  if(columnInfo === undefined) {
    myColumnInfo = Excel.generateDefaultColumnInfo(myPath, ",");
  }

  var excel = new Excel(myPath, Excel.IN_FILETYPE_CSV, myColumnInfo);
  excel.saveAsExcel(path + ".xls");
  excel.quit();

  // 一時ファイルを削除する
  if(path !== myPath) {
    File.unlink(myPath);
  }
}

/**
 * タブ区切りファイルをExcelファイルに変換する。
 * @param {String} path タブ区切りファイルのパスを文字列で指定します。
 * @param {Array} columnInfo (オプション)タブ区切りファイルのカラム情報を指定します。<br>
<a href = "Excel.html#.COLUMN_TYPE_GENERAL">カラムタイプ</a>参照(初期値 全列に対しConst.COLUMN_TYPE_TEXT)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertFromTsv = function(path, columnInfo) {
  var myColumnInfo = columnInfo;
  if(columnInfo === undefined) {
    myColumnInfo = Excel.generateDefaultColumnInfo(path, "\t");
  }

  var excel = new Excel(path, Excel.IN_FILETYPE_TSV, myColumnInfo);
  excel.saveAsExcel(path + ".xls");
  excel.quit();
}

// 1行目の内容に基づいてデフォルトのcolumnInfoを生成する。
Excel.generateDefaultColumnInfo = function(path, separateChar) {
  var columnInfo = [];

  File.open(path, "r", function(file) {
    var firstLine = file.readLine();
    var columnCount = firstLine.split(separateChar).length;

    var i;
    for(i=0; i<columnCount; i++) {
      columnInfo.push([(i+1), Excel.COLUMN_TYPE_TEXT]);
    }
  });

  return columnInfo;
}

// Prototypes of Excel
Excel.prototype = {
  /**
   * ワークブックに含まれる、ワークシートを処理します。
   * @param {Function} block ブロック
   */
  each: function(block) {
    var enumSheet = new Enumerator(this.workbookObj.Sheets);
    for (;!enumSheet.atEnd(); enumSheet.moveNext()) {
       block(new ExcelSheet(enumSheet.item()));
    }
  },
  /**
   * 指定したsheetNameのワークシートを取得します。
   * @param {String} sheetName 取得するワークシート名
   * @return {Object} ワークシートオブジェクト
   */
  getSheetByName: function(sheetName) {
    var result;
    this.each(function(sheet) {
      if(sheet.getName() === sheetName) {
        result = sheet;
      }
    });

    return result;
  },
  /**
   * 指定したindexのワークシートを取得します。
   * @param {Number} index 取得するワークシートのインデックス番号(0オリジン)
   * @return {ExcelSheet} ワークシートオブジェクト
   */
  getSheetByIndex: function(index) {
    var result;
    var i=0;
    this.each(function(sheet) {
      if(i === index) {
        result = sheet;
      }
      i++;
    });

    return result;
  },
  /**
   * 指定したindexのワークシートの後にシートを1枚追加します。
   * @param {String} sheetName 追加するシートのシート名
   * @param {Number} index (オプション)追加するシートの位置(初期値 最後尾)
   * @return {ExcelSheet} ワークシートオブジェクト
   */
  addSheet: function(sheetName, index) {
    var myIndex = index || this.getSheetCount() - 1;

    var targetSheet = this.getSheetByIndex(myIndex).getRawObject();
    var addedSheet = this.workbookObj.Sheets.Add(null, targetSheet, 1, null);
    addedSheet.Name = sheetName;

    return new ExcelSheet(addedSheet);
  },
  /**
   * ワークシート数を取得します。
   * @return {Number} ワークシート数
   */
  getSheetCount: function() {
    return this.workbookObj.Sheets.Count;
  },
  /** 
   * Excelファイルを保存します。
   */
  save: function() {
    this.workbookObj.Save();
  },
  getFullPathName: function(fileName) {
    var myFileName = fileName;
    if(myFileName.indexOf(File.PATH_SEPARATOR) === -1) {
      myFileName = File.realpath(myFileName);
    }

    return myFileName;
  },
  /** 
   * Excelファイルをファイル名を指定して保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAs: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName));
  },
  /** 
   * Excelファイルをファイル名を指定してCSV形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsCsv: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_CSV);
  },
  /** 
   * Excelファイルをファイル名を指定してタブ区切り形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsTsv: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_TSV);
  },
  /** 
   * Excelファイルをファイル名を指定してExcel形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsExcel: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_EXCEL);
  },
  /**
   * Excelファイルを保存せずに閉じます。
   */
  quitDiscardChanges: function() {
    this.workbookObj.Close(false);
    this.excelObj.Quit();
  },
  /** 
   * Excelを終了します。<br/>
   * 未保存の変更がある場合は、ダイアログが表示されます。
   */
  quit: function() {
    this.workbookObj.Close();
    this.excelObj.Quit();
  },
  /**
   * 先頭から指定した枚数のシートを削除します。
   * @param {Number} numSheet 削除するシート数
   */
  removeSheets: function(numSheets) {
    var i;
    for(i=0; i<numSheets; i++) {
      var sheet = this.getSheetByIndex(0);
      sheet.remove();
    }
  },
  /**
   * 先頭のシートを選択します。
   */
  activateHeadSheet: function() {
    this.getSheetByIndex(0).activate();
  }
};

/** 
 * Excelシートオブジェクトを作成します。
 * @class Excelシートの操作と読み書きを行うためのクラスです。<br/>
 * Excelがインストールされている必要があります。
 * @param {String} sheetObj Worksheetオブジェクトを指定します。
 */
var ExcelSheet = function(sheetObj) {
  this.sheetObj = sheetObj;
};

// Prototypes of ExcelSheet
ExcelSheet.prototype = {
  /**
   * ワークシート名を取得します。
   * @return {String} ワークシート名
   */
  getName: function() {
    return this.sheetObj.Name;
  },
  /**
   * ワークシート名を設定します。
   * @param {String} sheetName ワークシート名
   */
  setName: function(sheetName) {
    return this.sheetObj.Name = sheetName;
  },
  /**
   * セルに値を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Object} value 設定する値
   */
  setValue: function(row, col, value) {
    this.sheetObj.Cells(row, col).Value = value;
  },
  /**
   * セルに値を設定します。
   * @param {Number} row 行
   * @param {Array} value 設定する値
   * @param {Number} coloffset (オプション)列の開始位置(初期値 先頭列)
   */
  setArrayValue: function(row, value, coloffset) {
    if(!coloffset) {
      coloffset = 1;
    }
    var valueLength = value.length;
    for(var i=0; i<valueLength; i++) {
      this.sheetObj.Cells(row, i+coloffset).Value = value[i];
    }
  },
  /**
   * セルの値を取得します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @return {Object} セルの値
   */
  getValue: function(row, col) {
    return this.sheetObj.Cells(row, col).Value;
  },
  /**
   * ラップしている生のWorksheetオブジェクトを返します。<br/>
   * 通常使いません。
   */
  getRawObject: function() {
    return this.sheetObj;
  },
  /**
   * オートフィットを行います。
   */
  autoFit: function() {
    this.sheetObj.Columns("A:IV").AutoFit();
  },
  /**
   * Cell(1, 1)のカレントリージョンに対して罫線を引きます。
   */
  drawBorder: function() {
    var range = this.sheetObj.Cells(1, 1).CurrentRegion;
    for(var i=7; i<=10; i++) {
      range.Borders(i).LineStyle = 1;
    }
    if(range.Columns.Count > 1) {
        range.Borders(11).LineStyle = 1;
    }
    if(range.Rows.Count > 1) {
        range.Borders(12).LineStyle = 1;
    }
  },
  /**
   * セルの書式を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} format 書式
   */
  setFormat: function(row, col, format) {
    this.sheetObj.Cells(row, col).NumberFormatLocal = format;
  },
  /**
   * セルにコメントを追加します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} comment コメント
   * @param {Boolean} visible コメントを表示するかどうか
   */
  addComment: function(row, col, comment, visible) {
    this.sheetObj.Cells(row, col).AddComment();
    this.sheetObj.Cells(row, col).Comment.Visible = visible;
    this.sheetObj.Cells(row, col).Comment.Text(comment);
  },
  /**
   * シートをアクティブ化します。
   */
  activate: function() {
    this.sheetObj.Activate();
  },
  /**
   * シートを削除します。
   */
  remove: function() {
    this.sheetObj.Delete();
  },
  /**
   * ズームを設定します。
   * @param {Number} zoom 拡大率(整数)
   */
  setZoom: function(zoom) {
    // シートをアクテイブ化する
    this.activate();

    // 拡大率を設定する
    this.sheetObj.Application.ActiveWindow.Zoom = zoom;
  },
  /**
   * センターヘッダを設定します。
   * @param {String} header センターヘッダに設定する内容
   */
  setCenterHeader: function(header) {
    this.sheetObj.PageSetup.CenterHeader = header;
  },
  /**
   * レフトヘッダを設定します。
   * @param {String} header レフトヘッダに設定する内容
   */
  setLeftHeader: function(header) {
    this.sheetObj.PageSetup.LeftHeader = header;
  },
  /**
   * ライトヘッダを設定します。
   * @param {String} header ライトヘッダに設定する内容
   */
  setRightHeader: function(header) {
    this.sheetObj.PageSetup.RightHeader = header;
  },
  /**
   * センターフッタを設定します。
   * @param {String} footer センターフッタに設定する内容
   */
  setCenterFooter: function(footer) {
    this.sheetObj.PageSetup.CenterFooter = footer;
  },
  /**
   * レフトフッタを設定します。
   * @param {String} footer レフトフッタに設定する内容
   */
  setLeftFooter: function(footer) {
    this.sheetObj.PageSetup.LeftFooter = footer;
  },
  /**
   * ライトフッタを設定します。
   * @param {String} footer ライトフッタに設定する内容
   */
  setRightFooter: function(footer) {
    this.sheetObj.PageSetup.RightFooter = footer;
  },
  /**
   * セル値をクリップボードへコピーします。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  copy: function(row, col) {
    this.sheetObj.Cells(row, col).Copy();
  },
  /**
   * セルの背景色を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} colorIndex カラーインデックス
   */
  setBackgroundColor: function(row, col, colorIndex) {
    this.sheetObj.Cells(row, col).Interior.ColorIndex = colorIndex;
  },
  /**
   * セルのフォント色を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} colorIndex カラーインデックス
   */
  setFontColor: function(row, col, colorIndex) {
    this.sheetObj.Cells(row, col).Font.ColorIndex = colorIndex;
  },
  /**
   * セルのフォントを太字に設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isBold 太字にするかどうか
   */
  setFontBold: function(row, col, isBold) {
    this.sheetObj.Cells(row, col).Font.Bold = isBold;
  },
  /**
   * セルのフォントをイタリックに設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isItalic 太字にするかどうか
   */
  setFontItalic: function(row, col, isItalic) {
    this.sheetObj.Cells(row, col).Font.Italic = isItalic;
  },
  /**
   * セルのフォントに取消し線を設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isStrikethrough 取消し線を設定するかどうか
   */
  setFontStrikethrough: function(row, col, isStrikethrough) {
    this.sheetObj.Cells(row, col).Font.Strikethrough = isStrikethrough;
  },
  /**
   * セルのフォントに下線を設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isUnderline 下線を設定するかどうか
   */
  setFontUnderline: function(row, col, isUnderline) {
    this.sheetObj.Cells(row, col).Font.Underline = isUnderline;
  },
  /**
   * セルを選択します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  selectCell: function(row, col) {
    this.sheetObj.Cells(row, col).Select();
  },
  /**
   * セル範囲を選択します。
   * @param {Number} row1 行（始点）
   * @param {Number} col1 列（始点）
   * @param {Number} row2 行（終点）
   * @param {Number} col2 列（終点）
   */
  selectRange: function(row1, col1, row2, col2) {
    this.sheetObj.Range(this.sheetObj.Cells(row1, col1), this.sheetObj.Cells(row2, col2)).Select();
  },
  /**
   * セルにハイパーリンクを設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} address リンク先
   * @param {String} displayText (オプション)表示文字列(初期値 リンク先)
   */
  addHyperLink: function(row, col, address, displayText) {
    if(!displayText) {
      displayText = address;
    }
    var anchor = this.sheetObj.Cells(row, col);
    this.sheetObj.Hyperlinks.Add(anchor, address, "", address, displayText);
  },
  /**
   * セルに設定されたハイパーリンクを削除します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  deleteHyperLink: function(row, col) {
    this.sheetObj.Cells(row, col).Hyperlinks.Delete();
  },
  /**
   * 指定した列のカラム幅を設定します。
   * @param {Number} col 列(1オリジン)
   * @param {Number} columnWidth 列幅
   */
  setColumnWidth: function(col, columnWidth) {
    this.sheetObj.Columns(col).ColumnWidth = columnWidth;
  },
  /**
   * 指定したセルにオートフィルタを設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  setAutoFilter: function(row, col) {
    this.sheetObj.Cells(row, col).CurrentRegion.AutoFilter();
  }
};
/**
 * インスタンス化しません。
 * @class イベントエントリをログファイルに追加する機能を提供するクラス
<pre class = "code">
使用例：
LogEvent.success("hoge1");
LogEvent.error("hoge2");
LogEvent.warning("hoge3");
LogEvent.information("hoge4");
LogEvent.auditSuccess("hoge5");
LogEvent.auditFailure("hoge6");
</pre>
 */
var LogEvent = {};

/**
 * イベントエントリをログファイルにSUCCESSで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.success = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_SUCCESS, strMessage);
};

/**
 * イベントエントリをログファイルにERRORで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.error = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_ERROR, strMessage);
};

/**
 * イベントエントリをログファイルにWARNINGで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.warning = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_WARNING, strMessage);
};

/**
 * イベントエントリをログファイルにINFORMATIONで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.information = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_INFORMATION, strMessage);
};

/**
 * イベントエントリをログファイルにAUDIT_SUCCESSで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.auditSuccess = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_AUDIT_SUCCESS, strMessage);
};

/**
 * イベントエントリをログファイルにAUDIT_FAILUREで追加します。
 * @param {String} strMessage ログエントリのテキストです。 
 */
LogEvent.auditFailure = function(strMessage) {
  Const.WSHELL.LogEvent(Const.LOG_EVENT_TYPE_AUDIT_FAILURE, strMessage);
};
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
/**
 * インスタンス化しません。
 * @class JSONの解析(parse)と、文字列化(stringify)をサポートするクラス
<pre class = "code">
使用例：
JSON.parse("[1,2,3]");
JSON.stringify([1,2,3]);
</pre>
 */
var JSON = {};

/**
 * 指定した文字列をオブジェクトを解析してオブジェクトとして返します。
 * @param {String} str JSON文字列
 * @return {Object} 解析結果のオブジェクト
 */
JSON.parse = function(str) {
  return eval('(' + str + ')');
};

/**
 * 指定したオブジェクトを文字列として返します。
 * @param {String} obj オブジェクト
 * @return {Object} 文字列化したオブジェクト
 */
JSON.stringify = function(obj) {
  return obj.toString();
};
