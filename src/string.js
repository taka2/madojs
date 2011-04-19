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

/**
 * chが全角かどうか判定する。<br/>
 * 二文字以上の文字列でも、先頭の文字のみ判定します。
 * @return {Boolean} chが全角の場合true、そうでない場合falseを返す。
 */
String.prototype.isZenkaku = function() {
  // 先頭一文字のみチェックする
  var text = this.valueOf();
  var c = text.substring(0, 1);

  // escapeした結果%u始まり、かつ、半角カナでない場合全角と判定する
  return (escape(c).substring(0, 2) === "%u") &&
          !c.isHankakuKana();
}

/**
 * chが半角カナかどうか判定する。<br/>
 * 二文字以上の文字列でも、先頭の文字のみ判定します。
 * @return {Boolean} chが半角カナの場合true、そうでない場合falseを返す。
 */
String.prototype.isHankakuKana = function() {
  // 先頭一文字のみチェックする
  var text = this.valueOf();
  var c = text.substring(0, 1);
  var code = c.charCodeAt(0);

  return(code >= 0xFF71 && code <= 0xFF90);
}

/**
 * 全角文字を2バイト、半角文字を1バイトとして、指定位置の文字列を返します。
 * @param {Number} start 開始位置
 * @param {Number} end 終了位置
 * @return {String} 指定位置の文字列
 */
String.prototype.substringb = function(start, end) {
  var text = this.valueOf();
  var currentIndex = 0;
  var result = "";

  for(var i=0; i<text.length; i++) {
    var ch = text.charAt(i);
    var chSize = 1;
    if(ch.isZenkaku()) {
      chSize = 2;
    }
    if(currentIndex >= start && (currentIndex + chSize) <= end) {
      result += ch;
    }

    currentIndex += chSize;
  }

  return result;
}
