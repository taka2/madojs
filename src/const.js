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
