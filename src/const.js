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
