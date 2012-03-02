/**
 * インスタンス化しません。
 * @class JSONの解析(parse)と、文字列化(stringify)をサポートするクラス
<pre class = "code">
使用例：
JSON.parse("[1,2,3]");
JSON.stringify([1,2,3]);
</pre>
 */
var JSON = {};

/**
 * 指定した文字列をオブジェクトを解析してオブジェクトとして返します。
 * @param {String} str JSON文字列
 * @return {Object} 解析結果のオブジェクト
 */
JSON.parse = function(str) {
  return eval('(' + str + ')');
};

/**
 * 指定したオブジェクトを文字列として返します。
 * @param {String} obj オブジェクト
 * @return {Object} 文字列化したオブジェクト
 */
JSON.stringify = function(obj) {
  return obj.toString();
};
