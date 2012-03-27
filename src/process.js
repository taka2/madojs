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
