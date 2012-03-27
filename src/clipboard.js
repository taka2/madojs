/** 
 * 新しいクリップボードを作成する
 * @class クリップボードへの格納、および、取得を行うクラス。
<pre class = "code">
使用例：
// "日本語"という文字列をクリップボードへコピーして、ペーストします。
Clipboard.open(function(clip) {
  clip.set("日本語");
  sendKeys("^v");
});
</pre>
 */
var Clipboard = function() {
  if(Excel.available()) {
    this.clip = new ClipboardExcel();
  } else {
    this.clip = new ClipboardIE();
  }
};

/** 
 * 新しいクリップボードを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したクリップボードを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Clipboard.open = function(block) {
  if (!isFunction(block)) {
    return new Clipboard();
  }

  try {
    var clip = new Clipboard();
    block(clip);
  } finally {
    if(clip) {
      clip.close();
    }
  }
};

// Prototypes of Clipboard
Clipboard.prototype = {
  /**
   * クリップボードへ文字列を格納します。
   * @param {String} text 格納する文字列
   */
  set: function(text) {
    this.clip.set(text);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    return this.clip.get();
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this.clip.close();
  }
};
