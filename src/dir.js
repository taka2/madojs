/** 
 * pathの存在チェックを行った上で、ディレクトリオブジェクトを作成します。
 * @class ディレクトリの操作を行うためのクラスです。
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
