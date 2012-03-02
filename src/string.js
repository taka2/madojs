// オリジナルのtoStringをorgToStringとして保存
String.prototype.orgToString = String.prototype.toString;

/**
 * オブジェクトの値を表す文字列を返します。
 * @return {String} オブジェクトの値を表す文字列
 */
String.prototype.toString = function() {
  return '"' + this.valueOf() + '"';
};

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
 * 文字列がstrで始まるかどうか
 * @param {String} str テストする文字列
 * @return {Boolean} 文字列がstrで始まる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".startsWith("a"); // true
 */
String.prototype.startsWith = function(str) {
  var text = this.valueOf();
  return(text.match("^" + str) === null ? false : true);
};

/**
 * 文字列がstrで終わるかどうか
 * @param {String} str テストする文字列
 * @return {Boolean} 文字列がstrで終わる場合true、そうでない場合はfalseを返す。
 * @example 使用例：
"abc".endsWith("c"); // true
 */
String.prototype.endsWith = function(str) {
  var text = this.valueOf();
  return(text.match(str + "$") === null ? false : true);
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

  var textLength = text.length;
  for(var i=0; i<textLength; i++) {
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

/**
 * 指定文字を、指定文字数に達するまで左側にパディングします。
 * @param {String} ch 文字
 * @param {Number} length 文字数
 * @return {String} 指定位置の文字列
 * @example 使用例：
"12".lpad("0", 5); // "00012"
 */
String.prototype.lpad = function(ch, length) {
  var text = this.valueOf();

  // 文字列が指定文字数に既に達している場合はそのまま返す。
  if(text.length >= length) {
    return text;
  }

  var result = text;
  for(var i=0; i<length - text.length; i++) {
    result = ch + result;
  }
  return result;
}
