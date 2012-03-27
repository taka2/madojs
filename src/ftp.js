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
