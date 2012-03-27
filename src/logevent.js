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
