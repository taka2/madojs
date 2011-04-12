/** 
 * 文字列先頭と末尾のスペースをトリムした文字列を生成して返します。
 * @return {String} トリムした文字列
 * @example 使用例：
print("@" + "   	abc			".trim() + "@"); // "@abc@"
 */
String.prototype.trim = function() {
  var text = this.valueOf();
  return (text || "").replace( /^\s+|\s+$/g, "" );
};

/**
 * 文字列がchで始まるかどうか
 * @param {String} ch テストする文字
 * @return {Boolean} 文字列がchで始まる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".startsWith("a"); // true
 */
String.prototype.startsWith = function(ch) {
  var text = this.valueOf();
  return(text.substring(0, 1) === ch);
};

/**
 * 文字列がchで終わるかどうか
 * @param {String} ch テストする文字
 * @return {Boolean} 文字列がchで終わる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".endsWith("c"); // true
 */
String.prototype.endsWith = function(ch) {
  var text = this.valueOf();
  var textLength = text.length;
  return(text.substring(textLength - 1, textLength) === ch);
};
