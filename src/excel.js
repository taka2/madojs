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
 * @param {Boolean} bNew (オプション)新しくファイルを作成するかどうか(初期値 false)
 * @throws bNewがfalse、かつ、ファイルが存在しない場合にスローされます。
 */
var Excel = function(path, bNew) {
  this.path = path;

  this.excelObj = new ActiveXObject("Excel.Application");
  if(bNew) {
    this.workbookObj = this.excelObj.Workbooks.Add();
  } else {
    if(!File.exist(path)) {
      throw new Error(-1, "File Not Found: " + path);
    }

    this.workbookObj = this.excelObj.Workbooks.Open(path);
  }
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
    return new Excel(path, false);
  }

  try {
    var excel = new Excel(path, false);
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
    return new Excel(path, false);
  }

  try {
    var excel = new Excel(path, false);
    block(excel);
  } finally {
    if (excel != null) {
      excel.quitDiscardChanges();
    }
  }
};

/** 
 * Excelファイルを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.create = function(block) {
  if(!isFunction(block)) {
    return new Excel(path, true);
  }

  try {
    var excel = new Excel("", true);
    block(excel);
  } finally {
    if (excel != null) {
      excel.quit();
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
   * @param {Number} index (オプション)追加するシートの位置(初期値 最後尾)
   * @return {ExcelSheet} ワークシートオブジェクト
   */
  addSheet: function(sheetName, index) {
    var myIndex = index || this.getSheetCount() - 1;

    var targetSheet = this.getSheetByIndex(myIndex).getRawObject();
    var addedSheet = this.workbookObj.Sheets.Add(null, targetSheet, 1, null);
    addedSheet.Name = sheetName;

    return new ExcelSheet(addedSheet);
  },
  /**
   * ワークシート数を取得します。
   * @return {Number} ワークシート数
   */
  getSheetCount: function() {
    return this.workbookObj.Sheets.Count;
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
   * ワークシート名を設定します。
   * @param {String} sheetName ワークシート名
   */
  setName: function(sheetName) {
    return this.sheetObj.Name = sheetName;
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
  },
  /**
   * オートフィットを行います。
   */
  autoFit: function() {
    this.sheetObj.Columns("A:IV").AutoFit();
  },
  /**
   * カレントリージョンに対して罫線を引きます。
   */
  drawBorder: function() {
    var range = this.sheetObj.Cells(1, 1).CurrentRegion;
    var toIndex = 12;
    if(range.Rows.Count === 1) {
      // 1行しかない場合は12は引かない
      toIndex = 11;
    }
    for(var i=7; i<=toIndex; i++) {
      range.Borders(i).LineStyle = 1;
    }
  },
  /**
   * セルの書式を設定します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {String} format 書式
   */
  setFormat: function(x, y, format) {
    this.sheetObj.Cells(x, y).NumberFormatLocal = format;
  },
  /**
   * セルにコメントを追加します。
   * @param {Number} x x座標
   * @param {Number} y y座標
   * @param {String} comment コメント
   * @param {Boolean} visible コメントを表示するかどうか
   */
  addComment: function(x, y, comment, visible) {
    this.sheetObj.Cells(x, y).AddComment();
    this.sheetObj.Cells(x, y).Comment.Visible = visible;
    this.sheetObj.Cells(x, y).Comment.Text(comment);
  },
  /**
   * シートをアクティブ化します。
   */
  activate: function() {
    this.sheetObj.Activate();
  },
  /**
   * シートを削除します。
   */
  remove: function() {
    this.sheetObj.Delete();
  },
  /**
   * ズームを設定します。
   * @param {Number} zoom 拡大率(整数)
   */
  setZoom: function(zoom) {
    // シートをアクテイブ化する
    this.activate();

    // 拡大率を設定する
    this.sheetObj.Application.ActiveWindow.Zoom = zoom;
  },
  /**
   * センターヘッダを設定します。
   * @param {String} header センターヘッダに設定する内容
   */
  setCenterHeader: function(header) {
    this.sheetObj.PageSetup.CenterHeader = header;
  },
  /**
   * レフトヘッダを設定します。
   * @param {String} header レフトヘッダに設定する内容
   */
  setLeftHeader: function(header) {
    this.sheetObj.PageSetup.LeftHeader = header;
  },
  /**
   * ライトヘッダを設定します。
   * @param {String} header ライトヘッダに設定する内容
   */
  setRightHeader: function(header) {
    this.sheetObj.PageSetup.RightHeader = header;
  },
  /**
   * センターフッタを設定します。
   * @param {String} footer センターフッタに設定する内容
   */
  setCenterFooter: function(footer) {
    this.sheetObj.PageSetup.CenterFooter = footer;
  },
  /**
   * レフトフッタを設定します。
   * @param {String} footer レフトフッタに設定する内容
   */
  setLeftFooter: function(footer) {
    this.sheetObj.PageSetup.LeftFooter = footer;
  },
  /**
   * ライトフッタを設定します。
   * @param {String} footer ライトフッタに設定する内容
   */
  setRightFooter: function(footer) {
    this.sheetObj.PageSetup.RightFooter = footer;
  }
};
