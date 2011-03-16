/** 
 * Excelオブジェクトを作成します。
 * @class Excelファイルの操作と読み書きを行うためのクラスです。<br/>
 * Excelがインストールされている必要があります。
<pre class = "code">
使用例：
// abc.xlsを開いて、Sheet1の、1行1列目の値を'hoge'に書き換えつつ、上書き保存する。
Excel.open("abc.xls", function(excel) {
  var sheet = excel.getSheetByName("Sheet1");
  sheet.setValue(1, 1, 'hoge');
  excel.save();
});
</pre>
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @throws ファイルが存在しない場合にスローされます。
 */
var Excel = function(path) {
  this.path = path;

  if(!File.exist(path)) {
    throw new Error(-1, "File Not Found: " + path);
  }

  this.excelObj = new ActiveXObject("Excel.Application");
  this.workbookObj = this.excelObj.Workbooks.Open(path);
};

/** 
 * Excelファイルを開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.open = function(path, block) {
  if(!isFunction(block)) {
    return new Excel(path);
  }

  try {
    var excel = new Excel(path);
    block(excel);
  } finally {
    if (excel != null) {
      excel.quit();
    }
  }
};

/** 
 * Excelファイルを読み取り専用(最後に変更を破棄)で開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.openReadonly = function(path, block) {
  if(!isFunction(block)) {
    return new Excel(path);
  }

  try {
    var excel = new Excel(path);
    block(excel);
  } finally {
    if (excel != null) {
      excel.quitDiscardChanges();
    }
  }
};

// Prototypes of Excel
Excel.prototype = {
  /**
   * ワークブックに含まれる、ワークシートを処理します。
   * @param {Function} block ブロック
   */
  each: function(block) {
    var enumSheet = new Enumerator(this.workbookObj.Sheets);
    for (;!enumSheet.atEnd(); enumSheet.moveNext()) {
       block(new ExcelSheet(enumSheet.item()));
    }
  },
  /**
   * 指定したsheetNameのワークシートを取得します。
   * @param {String} sheetName 取得するワークシート名
   * @return {Object} ワークシートオブジェクト
   */
  getSheetByName: function(sheetName) {
    var result;
    this.each(function(sheet) {
      if(sheet.getName() === sheetName) {
        result = sheet;
      }
    });

    return result;
  },
  /**
   * 指定したindexのワークシートを取得します。
   * @param {Number} index 取得するワークシートのインデックス番号
   * @return {ExcelSheet} ワークシートオブジェクト
   */
  getSheetByIndex: function(index) {
    var result;
    var i=0;
    this.each(function(sheet) {
      if(i === index) {
        result = sheet;
      }
      i++;
    });

    return result;
  },
  /**
   * 指定したindexのワークシートの後にシートを1枚追加します。
   * @param {String} sheetName 追加するシートのシート名
   * @param {Number} index 追加するシートの位置
   * @return {ExcelSheet} ワークシートオブジェクト
   */
  addSheet: function(sheetName, index) {
    var targetSheet = this.getSheetByIndex(index).getRawObject();
    var addedSheet = this.workbookObj.Sheets.Add(null, targetSheet, 1, null);
    addedSheet.Name = sheetName;

    return new ExcelSheet(addedSheet);
  },
  /** 
   * Excelファイルを保存します。
   */
  save: function() {
    this.workbookObj.Save();
  },
  /** 
   * Excelファイルをファイル名を指定して保存します。
   */
  saveAs: function(fileName) {
    this.workbookObj.SaveAs(fileName);
  },
  /**
   * Excelファイルを保存せずに閉じます
   */
  quitDiscardChanges: function() {
    this.workbookObj.Close(false);
    this.excelObj.Quit();
  },
  /** 
   * Excelを終了します。<br/>
   * 未保存の変更がある場合は、ダイアログが表示されます。
   */
  quit: function() {
    this.workbookObj.Close();
    this.excelObj.Quit();
  }
};

/** 
 * Excelシートオブジェクトを作成します。
 * @class Excelシートの操作と読み書きを行うためのクラスです。<br/>
 * Excelがインストールされている必要があります。
 * @param {String} sheetObj Worksheetオブジェクトを指定します。
 */
var ExcelSheet = function(sheetObj) {
  this.sheetObj = sheetObj;
};

// Prototypes of ExcelSheet
ExcelSheet.prototype = {
  /**
   * ワークシート名を取得します。
   * @return {String} ワークシート名
   */
  getName: function() {
    return this.sheetObj.Name;
  },
  /**
   * セルに値を設定します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {Object} value 設定する値
   */
  setValue: function(x, y, value) {
    this.sheetObj.Cells(x, y).Value = value;
  },
  /**
   * セルの値を取得します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @return {Object} セルの値
   */
  getValue: function(x, y) {
    return this.sheetObj.Cells(x, y).Value;
  },
  /**
   * ラップしている生のWorksheetオブジェクトを返します。<br/>
   * 通常使いません。
   */
  getRawObject: function() {
    return this.sheetObj;
  }
};
