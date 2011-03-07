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
  },
  /**
   * Back Spaceをキー送信します。
   * @return this
   */
  sendBackspace: function(ch) {
    sendKeys("{BACKSPACE}");
    return this;
  },
  /**
   * Breakをキー送信します。
   * @return this
   */
  sendBreak: function(ch) {
    sendKeys("{BREAK}");
    return this;
  },
  /**
   * Caps Lockをキー送信します。
   * @return this
   */
  sendCapsLock: function(ch) {
    sendKeys("{CAPSLOCK}");
    return this;
  },
  /**
   * Deleteをキー送信します。
   * @return this
   */
  sendDelete: function(ch) {
    sendKeys("{DELETE}");
    return this;
  },
  /**
   * ↓をキー送信します。
   * @return this
   */
  sendDownArrow: function(ch) {
    sendKeys("{DOWN}");
    return this;
  },
  /**
   * Endをキー送信します。
   * @return this
   */
  sendEnd: function(ch) {
    sendKeys("{END}");
    return this;
  },
  /**
   * Enterをキー送信します。
   * @return this
   */
  sendEnter: function() {
    sendKeys("{ENTER}");
    return this;
  },
  /**
   * Escをキー送信します。
   * @return this
   */
  sendEsc: function(ch) {
    sendKeys("{ESC}");
    return this;
  },
  /**
   * Helpをキー送信します。
   * @return this
   */
  sendHelp: function(ch) {
    sendKeys("{HELP}");
    return this;
  },
  /**
   * Homeをキー送信します。
   * @return this
   */
  sendHome: function(ch) {
    sendKeys("{HOME}");
    return this;
  },
  /**
   * Insertをキー送信します。
   * @return this
   */
  sendInsert: function(ch) {
    sendKeys("{INSERT}");
    return this;
  },
  /**
   * ←をキー送信します。
   * @return this
   */
  sendLeftArrow: function(ch) {
    sendKeys("{LEFT}");
    return this;
  },
  /**
   * NumLockをキー送信します。
   * @return this
   */
  sendNumLock: function(ch) {
    sendKeys("{NUMLOCK}");
    return this;
  },
  /**
   * Page Downをキー送信します。
   * @return this
   */
  sendPageDown: function(ch) {
    sendKeys("{PGDN}");
    return this;
  },
  /**
   * Page Upをキー送信します。
   * @return this
   */
  sendPageUp: function(ch) {
    sendKeys("{PGUP}");
    return this;
  },
  /**
   * Print Screenをキー送信します。
   * @return this
   */
  sendPrintScreen: function(ch) {
    sendKeys("{PRTSC}");
    return this;
  },
  /**
   * →をキー送信します。
   * @return this
   */
  sendRightArrow: function(ch) {
    sendKeys("{RIGHT}");
    return this;
  },
  /**
   * Scroll Lockをキー送信します。
   * @return this
   */
  sendScrollLock: function(ch) {
    sendKeys("{SCROLLLOCK}");
    return this;
  },
  /**
   * Tabをキー送信します。
   * @return this
   */
  sendTab: function() {
    sendKeys("{TAB}");
    return this;
  },
  /**
   * ↑をキー送信します。
   * @return this
   */
  sendUpArrow: function() {
    sendKeys("{UP}");
    return this;
  },
  /**
   * F1をキー送信します。
   * @return this
   */
  sendF1: function() {
    sendKeys("{F1}");
    return this;
  },
  /**
   * F2をキー送信します。
   * @return this
   */
  sendF2: function() {
    sendKeys("{F2}");
    return this;
  },
  /**
   * F3をキー送信します。
   * @return this
   */
  sendF3: function() {
    sendKeys("{F3}");
    return this;
  },
  /**
   * F4をキー送信します。
   * @return this
   */
  sendF4: function() {
    sendKeys("{F4}");
    return this;
  },
  /**
   * F5をキー送信します。
   * @return this
   */
  sendF5: function() {
    sendKeys("{F5}");
    return this;
  },
  /**
   * F6をキー送信します。
   * @return this
   */
  sendF6: function() {
    sendKeys("{F6}");
    return this;
  },
  /**
   * F7をキー送信します。
   * @return this
   */
  sendF7: function() {
    sendKeys("{F7}");
    return this;
  },
  /**
   * F8をキー送信します。
   * @return this
   */
  sendF8: function() {
    sendKeys("{F8}");
    return this;
  },
  /**
   * F9をキー送信します。
   * @return this
   */
  sendF9: function() {
    sendKeys("{F9}");
    return this;
  },
  /**
   * F10をキー送信します。
   * @return this
   */
  sendF10: function() {
    sendKeys("{F10}");
    return this;
  },
  /**
   * F11をキー送信します。
   * @return this
   */
  sendF11: function() {
    sendKeys("{F11}");
    return this;
  },
  /**
   * F12をキー送信します。
   * @return this
   */
  sendF12: function() {
    sendKeys("{F12}");
    return this;
  },
  /**
   * F13をキー送信します。
   * @return this
   */
  sendF13: function() {
    sendKeys("{F13}");
    return this;
  },
  /**
   * F14をキー送信します。
   * @return this
   */
  sendF14: function() {
    sendKeys("{F14}");
    return this;
  },
  /**
   * F15をキー送信します。
   * @return this
   */
  sendF15: function() {
    sendKeys("{F15}");
    return this;
  },
  /**
   * F16をキー送信します。
   * @return this
   */
  sendF16: function() {
    sendKeys("{F16}");
    return this;
  }
};
