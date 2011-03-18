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
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeys: function(text, num) {
    sendKeys(text, num);
    return this;
  },
  /**
   * 指定した日本語textをキー送信します。
   * @param {String} text 送信する文字列
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendJapaneseKeys: function(text, num) {
    Clipboard.open(function(clip) {
      for(var i=0; i<num; i++) {
        clip.set(text);
        sendKeys("^v");
      }
    });
    return this;
  },
  /**
   * Shiftを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithShift: function(ch, num) {
    sendKeys("+" + ch, num);
    return this;
  },
  /**
   * Controlを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithControl: function(ch, num) {
    sendKeys("^" + ch, num);
    return this;
  },
  /**
   * Altを押しながら指定したchをキー送信します。
   * @param {String} ch 送信する文字
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendKeyWithAlt: function(ch, num) {
    sendKeys("%" + ch, num);
    return this;
  },
  /**
   * Back Spaceをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBackspace: function(ch, num) {
    sendKeys("{BACKSPACE}", num);
    return this;
  },
  /**
   * Breakをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendBreak: function(ch, num) {
    sendKeys("{BREAK}", num);
    return this;
  },
  /**
   * Caps Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendCapsLock: function(ch, num) {
    sendKeys("{CAPSLOCK}", num);
    return this;
  },
  /**
   * Deleteをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDelete: function(ch, num) {
    sendKeys("{DELETE}", num);
    return this;
  },
  /**
   * ↓をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendDownArrow: function(ch, num) {
    sendKeys("{DOWN}", num);
    return this;
  },
  /**
   * Endをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnd: function(ch, num) {
    sendKeys("{END}", num);
    return this;
  },
  /**
   * Enterをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEnter: function(num) {
    sendKeys("{ENTER}", num);
    return this;
  },
  /**
   * Escをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendEsc: function(num) {
    sendKeys("{ESC}", num);
    return this;
  },
  /**
   * Helpをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHelp: function(num) {
    sendKeys("{HELP}", num);
    return this;
  },
  /**
   * Homeをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendHome: function(num) {
    sendKeys("{HOME}", num);
    return this;
  },
  /**
   * Insertをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendInsert: function(num) {
    sendKeys("{INSERT}", num);
    return this;
  },
  /**
   * ←をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendLeftArrow: function(num) {
    sendKeys("{LEFT}", num);
    return this;
  },
  /**
   * NumLockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendNumLock: function(num) {
    sendKeys("{NUMLOCK}", num);
    return this;
  },
  /**
   * Page Downをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageDown: function(num) {
    sendKeys("{PGDN}", num);
    return this;
  },
  /**
   * Page Upをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPageUp: function(num) {
    sendKeys("{PGUP}", num);
    return this;
  },
  /**
   * Print Screenをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendPrintScreen: function(num) {
    sendKeys("{PRTSC}", num);
    return this;
  },
  /**
   * →をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendRightArrow: function(num) {
    sendKeys("{RIGHT}", num);
    return this;
  },
  /**
   * Scroll Lockをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendScrollLock: function(num) {
    sendKeys("{SCROLLLOCK}", num);
    return this;
  },
  /**
   * Tabをキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendTab: function(num) {
    sendKeys("{TAB}", num);
    return this;
  },
  /**
   * ↑をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendUpArrow: function(num) {
    sendKeys("{UP}", num);
    return this;
  },
  /**
   * F1をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF1: function(num) {
    sendKeys("{F1}", num);
    return this;
  },
  /**
   * F2をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF2: function(num) {
    sendKeys("{F2}", num);
    return this;
  },
  /**
   * F3をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF3: function(num) {
    sendKeys("{F3}", num);
    return this;
  },
  /**
   * F4をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF4: function(num) {
    sendKeys("{F4}", num);
    return this;
  },
  /**
   * F5をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF5: function(num) {
    sendKeys("{F5}", num);
    return this;
  },
  /**
   * F6をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF6: function(num) {
    sendKeys("{F6}", num);
    return this;
  },
  /**
   * F7をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF7: function(num) {
    sendKeys("{F7}", num);
    return this;
  },
  /**
   * F8をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF8: function(num) {
    sendKeys("{F8}", num);
    return this;
  },
  /**
   * F9をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF9: function(num) {
    sendKeys("{F9}", num);
    return this;
  },
  /**
   * F10をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF10: function(num) {
    sendKeys("{F10}", num);
    return this;
  },
  /**
   * F11をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF11: function(num) {
    sendKeys("{F11}", num);
    return this;
  },
  /**
   * F12をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF12: function(num) {
    sendKeys("{F12}", num);
    return this;
  },
  /**
   * F13をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF13: function(num) {
    sendKeys("{F13}", num);
    return this;
  },
  /**
   * F14をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF14: function(num) {
    sendKeys("{F14}", num);
    return this;
  },
  /**
   * F15をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF15: function(num) {
    sendKeys("{F15}", num);
    return this;
  },
  /**
   * F16をキー送信します。
   * @param {Number} num (オプション)送信する回数(初期値 1)
   * @return this
   */
  sendF16: function(num) {
    sendKeys("{F16}", num);
    return this;
  }
};
