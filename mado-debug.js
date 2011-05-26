/**
 * インスタンス化しません。
 * @class MadoJSの定数を定義するクラス。
 */
var Const = {};

// Static members of Global
Const.FSO = new ActiveXObject("Scripting.FileSystemObject");
Const.WSHELL = WScript.CreateObject("WScript.Shell");
Const.SHELLAPP = WScript.CreateObject("Shell.Application");

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
 * 引数で構成された配列
 */
var ARGV = [];
(function() {
  var argvLength = WScript.Arguments.length;
  for(var i=0; i<argvLength; i++) {
    ARGV.push(WScript.Arguments(i));
  }
}());

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
 */
var sendKeys = function (key, num) {
  var myNum = num || 1;
  for(var i=0; i<myNum; i++) {
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
/** 
 * 各要素に対してブロックを評価します。
 * @param {Function} block ブロック
 * @example 使用例：
 * var x = [1, 2, 3].each(function(i) {
 *   return i*2;
 * });
 *
 * print(x); // 2,4,6
 */
Array.prototype.each = function(block) {
  var result = [];

  var thisLength = this.length;
  for(var i=0; i<thisLength; i++) {
    result.push(block(this[i]));
  }

  return result;
};

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
  return yyyy + mm + dd;
};

/** 
 * 日付をYYYYMMDDHH24MISS形式で返します。
 * @return {String} YYYYMMDDHH24MISS形式の日付
 * @example 使用例：
print(new Date().getYYYYMMDDHH24MISS());
 */
Date.prototype.getYYYYMMDDHH24MISS = function() {
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
  return yyyy + mm + dd + hh + mi + ss;
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
 * 文字列がchで始まるかどうか
 * @param {String} ch テストする文字
 * @return {Boolean} 文字列がchで始まる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".startsWith("a"); // true
 */
String.prototype.startsWith = function(ch) {
  var text = this.valueOf();
  return(text.substring(0, 1) === ch);
};

/**
 * 文字列がchで終わるかどうか
 * @param {String} ch テストする文字
 * @return {Boolean} 文字列がchで終わる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".endsWith("c"); // true
 */
String.prototype.endsWith = function(ch) {
  var text = this.valueOf();
  var textLength = text.length;
  return(text.substring(textLength - 1, textLength) === ch);
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
  if(!basedir) {
    basedir = Dir.getwd();
  }

  Dir.cd(basedir, function(path) {
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
   * ファイルへ一行書き込みます。
   * @param {String} line 書きこむ文字列
   */
  puts: function(line) {
    this.ts.WriteLine(line);
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
 * @return {String} レスポンステキスト
 * @example 使用例：
print(HTTP.get("www.yahoo.co.jp", "/index.html", 80));
 */
HTTP.get = function(address, path, port) {
  var myPort = port || DEFAULT_PORT_NUMBER;
  var http = new HTTP(address, port);
  return http.get(path);
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行い、printします。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @example 使用例：
HTTP.get_print("www.yahoo.co.jp", "/index.html", 80);
 */
HTTP.get_print = function(address, path, port) {
  print(HTTP.get(address, path, port));
};

// Prototypes of HTTP
HTTP.prototype = {
  /** 
   * 指定されたpath、headerでGETリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  get: function(path, header, block) {
    return this.request("GET", path, header, "", block);
  },
  /** 
   * 指定されたpath、headerでPOSTリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ(初期値 [])
   * @param {Object} body (オプション)HTTPボディ(初期値 undefined)
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  post: function(path, header, body, block) {
    return this.request("POST", path, header, body, block);
  },
  // private
  request: function(method, path, header, body, block) {
    // パラメータの調整
    var myHeader = header || [];

    // リクエストの準備
    var xhr = new ActiveXObject("Microsoft.XMLHTTP");
    xhr.open(method, "http://" + this.host + ":" + this.port + path, false);

    for(var headerName in myHeader) {
      xhr.setRequestHeader(headerName, myHeader[headerName]);
    }

    // リクエストの送信
    xhr.send(body);

    // レスポンスの処理
    if(xhr.readyState === 4) {
      if(xhr.status === 200) {
        if(isFunction(block)) {
          block(xhr.responseText);
        } else {
          return xhr.responseText;
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
 */
var KeySender = function(targetWindow) {
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
    sendKeys(text, num);
    return this;
  },
  /**
   * 指定した日本語textをキー送信します。
   * @param {String} text 送信する文字列
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendJapaneseKeys: function(text, num) {
    Clipboard.open(function(clip) {
      for(var i=0; i<num; i++) {
        clip.set(text);
        sendKeys("^v");
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
    sendKeys("+" + ch, num);
    return this;
  },
  /**
   * Controlを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithControl: function(ch, num) {
    sendKeys("^" + ch, num);
    return this;
  },
  /**
   * Altを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithAlt: function(ch, num) {
    sendKeys("%" + ch, num);
    return this;
  },
  /**
   * Back Spaceをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBackspace: function(ch, num) {
    sendKeys("{BACKSPACE}", num);
    return this;
  },
  /**
   * Breakをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBreak: function(ch, num) {
    sendKeys("{BREAK}", num);
    return this;
  },
  /**
   * Caps Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendCapsLock: function(ch, num) {
    sendKeys("{CAPSLOCK}", num);
    return this;
  },
  /**
   * Deleteをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDelete: function(ch, num) {
    sendKeys("{DELETE}", num);
    return this;
  },
  /**
   * ↓をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDownArrow: function(ch, num) {
    sendKeys("{DOWN}", num);
    return this;
  },
  /**
   * Endをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnd: function(ch, num) {
    sendKeys("{END}", num);
    return this;
  },
  /**
   * Enterをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnter: function(num) {
    sendKeys("{ENTER}", num);
    return this;
  },
  /**
   * Escをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEsc: function(num) {
    sendKeys("{ESC}", num);
    return this;
  },
  /**
   * Helpをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHelp: function(num) {
    sendKeys("{HELP}", num);
    return this;
  },
  /**
   * Homeをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHome: function(num) {
    sendKeys("{HOME}", num);
    return this;
  },
  /**
   * Insertをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendInsert: function(num) {
    sendKeys("{INSERT}", num);
    return this;
  },
  /**
   * ←をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendLeftArrow: function(num) {
    sendKeys("{LEFT}", num);
    return this;
  },
  /**
   * NumLockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendNumLock: function(num) {
    sendKeys("{NUMLOCK}", num);
    return this;
  },
  /**
   * Page Downをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageDown: function(num) {
    sendKeys("{PGDN}", num);
    return this;
  },
  /**
   * Page Upをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageUp: function(num) {
    sendKeys("{PGUP}", num);
    return this;
  },
  /**
   * Print Screenをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPrintScreen: function(num) {
    sendKeys("{PRTSC}", num);
    return this;
  },
  /**
   * →をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendRightArrow: function(num) {
    sendKeys("{RIGHT}", num);
    return this;
  },
  /**
   * Scroll Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendScrollLock: function(num) {
    sendKeys("{SCROLLLOCK}", num);
    return this;
  },
  /**
   * Tabをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendTab: function(num) {
    sendKeys("{TAB}", num);
    return this;
  },
  /**
   * ↑をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendUpArrow: function(num) {
    sendKeys("{UP}", num);
    return this;
  },
  /**
   * F1をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF1: function(num) {
    sendKeys("{F1}", num);
    return this;
  },
  /**
   * F2をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF2: function(num) {
    sendKeys("{F2}", num);
    return this;
  },
  /**
   * F3をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF3: function(num) {
    sendKeys("{F3}", num);
    return this;
  },
  /**
   * F4をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF4: function(num) {
    sendKeys("{F4}", num);
    return this;
  },
  /**
   * F5をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF5: function(num) {
    sendKeys("{F5}", num);
    return this;
  },
  /**
   * F6をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF6: function(num) {
    sendKeys("{F6}", num);
    return this;
  },
  /**
   * F7をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF7: function(num) {
    sendKeys("{F7}", num);
    return this;
  },
  /**
   * F8をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF8: function(num) {
    sendKeys("{F8}", num);
    return this;
  },
  /**
   * F9をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF9: function(num) {
    sendKeys("{F9}", num);
    return this;
  },
  /**
   * F10をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF10: function(num) {
    sendKeys("{F10}", num);
    return this;
  },
  /**
   * F11をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF11: function(num) {
    sendKeys("{F11}", num);
    return this;
  },
  /**
   * F12をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF12: function(num) {
    sendKeys("{F12}", num);
    return this;
  },
  /**
   * F13をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF13: function(num) {
    sendKeys("{F13}", num);
    return this;
  },
  /**
   * F14をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF14: function(num) {
    sendKeys("{F14}", num);
    return this;
  },
  /**
   * F15をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF15: function(num) {
    sendKeys("{F15}", num);
    return this;
  },
  /**
   * F16をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF16: function(num) {
    sendKeys("{F16}", num);
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
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Boolean} bNew (オプション)新しくファイルを作成するかどうか(初期値 false)
 * @throws bNewがfalse、かつ、ファイルが存在しない場合にスローされます。
 */
var Excel = function(path, bNew) {
  this.path = path;

  this.excelObj = new ActiveXObject("Excel.Application");
  if(bNew) {
    this.workbookObj = this.excelObj.Workbooks.Add();
  } else {
    if(!File.exist(path)) {
      throw new Error(-1, "File Not Found: " + path);
    }

    this.workbookObj = this.excelObj.Workbooks.Open(path);
  }
};

/** 
 * Excelファイルを開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.open = function(path, block) {
  if(!isFunction(block)) {
    return new Excel(path, false);
  }

  try {
    var excel = new Excel(path, false);
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
    return new Excel(path, false);
  }

  try {
    var excel = new Excel(path, false);
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
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.create = function(block) {
  if(!isFunction(block)) {
    return new Excel("", true);
  }

  try {
    var excel = new Excel("", true);
    block(excel);
  } finally {
    if(excel) {
      excel.quit();
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
   * @param {Number} index 取得するワークシートのインデックス番号
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
  /** 
   * Excelファイルをファイル名を指定して保存します。
   */
  saveAs: function(fileName) {
    this.workbookObj.SaveAs(fileName);
  },
  /**
   * Excelファイルを保存せずに閉じます
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
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {Object} value 設定する値
   */
  setValue: function(x, y, value) {
    this.sheetObj.Cells(x, y).Value = value;
  },
  /**
   * セルの値を取得します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @return {Object} セルの値
   */
  getValue: function(x, y) {
    return this.sheetObj.Cells(x, y).Value;
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
   * カレントリージョンに対して罫線を引きます。
   */
  drawBorder: function() {
    var range = this.sheetObj.Cells(1, 1).CurrentRegion;
    var toIndex = 12;
    if(range.Rows.Count === 1) {
      // 1行しかない場合は12は引かない
      toIndex = 11;
    }
    for(var i=7; i<=toIndex; i++) {
      range.Borders(i).LineStyle = 1;
    }
  },
  /**
   * セルの書式を設定します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {String} format 書式
   */
  setFormat: function(x, y, format) {
    this.sheetObj.Cells(x, y).NumberFormatLocal = format;
  },
  /**
   * セルにコメントを追加します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {String} comment コメント
   * @param {Boolean} visible コメントを表示するかどうか
   */
  addComment: function(x, y, comment, visible) {
    this.sheetObj.Cells(x, y).AddComment();
    this.sheetObj.Cells(x, y).Comment.Visible = visible;
    this.sheetObj.Cells(x, y).Comment.Text(comment);
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
   * @param {Number} x x座標
   * @param {Number} y y座標
   */
  copy: function(x, y) {
    this.sheetObj.Cells(x, y).Copy();
  }
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
AdoAccessConnection.prototype = AdoConnection.prototype;
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
AdoOracleConnection.prototype = AdoConnection.prototype;
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
