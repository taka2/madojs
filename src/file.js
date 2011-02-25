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
