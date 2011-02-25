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
