/**
 * インスタンス化しません。
 * @class ADODB.Streamのラッパクラス
<pre class = "code">
使用例：
// responseBodyをShift_JISにエンコーディング
print(Stream.binaryToText(responseBody, "Shift_JIS"));
</pre>
 */
var Stream = {};

/**
 * 文字エンコーディング：自動識別
 */
Stream.CHARACTER_ENCODING_AUTODETECT = "_autodetect";

/**
 * バイナリデータを指定したエンコーディングのテキストに変換します。
 * @param {Object} data バイナリデータ
 * @param {String} encoding 変換先のエンコーディング
 * @return {String} エンコーディングしたテキスト
 */
Stream.binaryToText = function(data, encoding) {
  var adodbstream = new ActiveXObject("ADODB.Stream");
  // Binaryモードにセット
  adodbstream.Type = 1;
  adodbstream.Open();
  adodbstream.Write(data);

  adodbstream.Position = 0;
  // Textモードにセット
  adodbstream.Type = 2;
  adodbstream.Charset = encoding;
  var text = adodbstream.ReadText();
  adodbstream.Close();

  return text;
};
