/**
 * インスタンス化しません。
 * @class MadoJSの定数を定義するクラス。
 */
var Const = {};

// Static members of Global
Const.FSO = new ActiveXObject("Scripting.FileSystemObject");
Const.WSHELL = WScript.CreateObject("WScript.Shell");

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
 */
var sendKeys = function (key) {
  Const.WSHELL.Sendkeys(key);
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
 * @param {Number} timeout タイムアウトを秒で指定
 */
var osShutdown = function (timeout) {
  if (timeout === null) timeout = 0;
  Process.exec("shutdown", ["-s", "-t", timeout], 0, false);
};

/**
 * OSを再起動します。
 * @param {Number} timeout タイムアウトを秒で指定
 * @param {String} option "-f"を指定すると、実行中のプロセスを警告なしで閉じます
 */
var osReboot = function (timeout, option) {
  var params = ["-r", "-t"];
  if (timeout === null) timeout = 0;
  params.push(timeout);

  if (option !== null) {
    params.push(option);
  }

  Process.exec("shutdown", params, 0, false);
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

  for(var i=0; i<this.length; i++) {
    result.push(block(this[i]));
  }

  return result;
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
      if(block) {
        block(item);
      } else {
        result.push(item);
      }
    }
  };

  dir.each(function(item) {
    if(Dir.exist(item)) {
      var subresult = Dir.find(item, pattern, block);
      for(var i=0; i<subresult.length; i++) {
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
    // パラメータの調整
    if(header === undefined) {
      header = [];
    }

    // リクエストの準備
    var xhr = new ActiveXObject("Microsoft.XMLHTTP");
    xhr.open("GET", "http://" + this.host + ":" + this.port + path, false);

    for(var headerName in header) {
      xhr.setRequestHeader(headerName, header[headerName]);
    }

    // リクエストの送信
    xhr.send();

    // レスポンスの処理
    if(xhr.readyState == 4) {
      if(xhr.status == 200) {
        if(block) {
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
Clipboard.open = function(block) {
  if (!isFunction(block)) {
    return new Clipboard();
  }

  try {
    var clip = new Clipboard();
    block(clip);
  } finally {
    if (clip != null) {
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
};

// Prototypes of Clipboard
KeySender.prototype = {
  /**
   * 指定したtextをキー送信します。
   * @param {String} text 送信する文字列
   * @return this
   */
  sendKeys: function(text) {
    sendKeys(text);
    return this;
  },
  /**
   * 指定した日本語textをキー送信します。
   * @param {String} text 送信する文字列
   * @return this
   */
  sendJapaneseKeys: function(text) {
    Clipboard.open(function(clip) {
      clip.set(text);
      sendKeys("^v");
    });
    return this;
  },
  /**
   * タブ文字をキー送信します。
   * @return this
   */
  sendTab: function() {
    sendKeys("{TAB}");
    return this;
  },
  /**
   * エンターをキー送信します。
   * @return this
   */
  sendEnter: function() {
    sendKeys("{ENTER}");
    return this;
  },
  /**
   * Shiftを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithShift: function(ch) {
    sendKeys("+" + ch);
    return this;
  },
  /**
   * Controlを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithControl: function(ch) {
    sendKeys("^" + ch);
    return this;
  },
  /**
   * Altを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithAlt: function(ch) {
    sendKeys("%" + ch);
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
 * @param {Array} args 引数
 * @param {Number} windowStyle <a href = "Const.html#.WINDOW_STYLE_MAX">ウインドウスタイル</a>参照
 * @param {Boolean} waitOnReturn trueの場合コマンドの実行が終わるまで待ちます。falseの場合は待ちません。
 */
Process.exec = function(command, args, windowStyle, waitOnReturn) {
  if(!args) {
    args = [];
  }
  if(!windowStyle) {
    windowStyle = Const.WINDOW_STYLE_NORMAL;
  }
  if(waitOnReturn === undefined) {
    waitOnReturn = true;
  }
  var commandLine = command + " " + args.join(" ");
  Const.WSHELL.Run(commandLine, windowStyle, waitOnReturn);
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
 * @throws ファイルが存在しない場合にスローされます。
 */
var Excel = function(path) {
  this.path = path;

  if(!File.exist(path)) {
    throw new Error(-1, "File Not Found: " + path);
  }

  this.excelObj = new ActiveXObject("Excel.Application");
  this.workbookObj = this.excelObj.Workbooks.Open(path);
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
    return new Excel(path);
  }

  try {
    var excel = new Excel(path);
    block(excel);
  } finally {
    if (excel != null) {
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
    if (excel != null) {
      excel.quitDiscardChanges();
    }
  }
};

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
  }
};
/** 
 * 新しいADOのOracleコネクションを作成する
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
AdoOracleConnection = function(dsName, userName, password) {
  this.con = new ActiveXObject("ADODB.Connection");

  // Open Connection
  this.con.Open("Provider=MSDAORA;Data Source=" + dsName + ";", userName, password);
};

/** 
 * 新しいADOのOracleコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのOracleコネクションを返します。
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのOracleコネクション
 */
AdoOracleConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOracleConnection(dsName, userName, password);
  }

  try {
    var oracon = new AdoOracleConnection(dsName, userName, password);
    block(oracon);
  } finally {
    if (oracon != null) {
      oracon.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOracleConnection.prototype = {
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql) {
    try {
      // SQLの実行
      var rs = this.con.Execute(sql);

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
    } catch(e) {
      throw e;
    } finally {
      if(rs) {
        rs.Close();
      }
    }
  },
  close: function() {
    this.con.Close();
  }
};
