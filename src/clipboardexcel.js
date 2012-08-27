/** 
 * 新しいクリップボードを作成する
 * @class Excelを利用した、クリップボードへの格納、および、取得を行うクラス。
 */
var ClipboardExcel = function() {
  this.excel = Excel.create();
  this.sheet = this.excel.getSheetByIndex(0);
  this.sheet.setFormat(1, 1, "@");
};

/** 
 * 新しいクリップボードを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したクリップボードを返します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
ClipboardExcel.open = function(block) {
  if (!isFunction(block)) {
    return new ClipboardExcel();
  }

  try {
    var clip = new ClipboardExcel();
    block(clip);
  } finally {
    if(clip) {
      clip.close();
    }
  }
};

// Prototypes of ClipboardExcel
ClipboardExcel.prototype = {
  /**
   * クリップボードへ文字列を格納します。
   * @param {String} text 格納する文字列
   */
  set: function(text) {
    this.sheet.setValue(1, 1, text);
    this.sheet.copy(1, 1);
  },
  /**
   * クリップボードの文字列を取得します。
   * @return {String} 取得した文字列
   */
  get: function() {
    return this.sheet.getValue(1, 1);
  },
  /**
   * クリップボードをクローズします。
   */
  close: function() {
    this.excel.quitDiscardChanges();
  }
};
