/** 
 * 新しいクリップボードを作成する
 * @class クリップボードへの格納、および、取得を行うクラス。
 */
var Clipboard = function() {
  this._ie = new ActiveXObject('InternetExplorer.Application');
  this._ie.navigate("about:blank");
  while(this._ie.Busy) {
    sleep(10);
  }
  this._ie.Visible = false;
  this._textarea = this._ie.document.createElement("textarea");
  this._ie.document.body.appendChild(this._textarea);
  this._textarea.focus();
  this._closed = false;
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
    if (clip != null) {
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
    if (this._closed) {
      return;
    }
    this._textarea.innerText = text;
    this._ie.execWB(17 /* select all */, 0);
    this._ie.execWB(12 /* copy */, 0);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    if (this._closed) {
      return;
    }
    this._textarea.innerText = "";
    this._ie.execWB(13 /* paste */, 0);
    return this._textarea.innerText;
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this._ie.Quit();
    this._closed = true;
  }
};
