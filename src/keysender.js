/** 
 * 新しいキー送信クラスを作成する。
 * @class 連続で自動的にキー送信を行う
<pre class = "code">
使用例：
// ノートパッドを起動
Process.exec("notepad", [], Const.WINDOW_STYLE_NORMAL, false);

// 起動を待つ
sleep(100);

// "日本語"という文字列を書きこんで、abc.txtに保存。
var ks = new KeySender("無題")
    .sendJapaneseKeys("日本語")
    .sendKeyWithControl("s")
    .sendKeys("abc.txt")
    .sendEnter();
</pre>
 */
var KeySender = function(targetWindow) {
  Const.WSHELL.AppActivate(targetWindow);
};

// Prototypes of Clipboard
KeySender.prototype = {
  /**
   * 指定したtextをキー送信します。
   * @param {String} text 送信する文字列
   * @return this
   */
  sendKeys: function(text) {
    sendKeys(text);
    return this;
  },
  /**
   * 指定した日本語textをキー送信します。
   * @param {String} text 送信する文字列
   * @return this
   */
  sendJapaneseKeys: function(text) {
    Clipboard.open(function(clip) {
      clip.set(text);
      sendKeys("^v");
    });
    return this;
  },
  /**
   * タブ文字をキー送信します。
   * @return this
   */
  sendTab: function() {
    sendKeys("{TAB}");
    return this;
  },
  /**
   * エンターをキー送信します。
   * @return this
   */
  sendEnter: function() {
    sendKeys("{ENTER}");
    return this;
  },
  /**
   * Shiftを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithShift: function(ch) {
    sendKeys("+" + ch);
    return this;
  },
  /**
   * Controlを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithControl: function(ch) {
    sendKeys("^" + ch);
    return this;
  },
  /**
   * Altを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @return this
   */
  sendKeyWithAlt: function(ch) {
    sendKeys("%" + ch);
    return this;
  }
};
