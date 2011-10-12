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
 * @param {String} path Excelファイルのパスを文字列で指定します。指定しない場合は新しくブックを作成します。
 * @param {Boolean} fileType (オプション)Excel以外のファイルを開くときに、指定します。
 * @param {Boolean} columnInfo (オプション)Excel以外のファイルを開くときに、カラムの情報を指定します。
 * @throws pathがundefinedでない、かつ、ファイルが存在しない場合にスローされます。
 */
var Excel = function(path, fileType, columnInfo) {
  this.path = path;
  this.excelObj = new ActiveXObject("Excel.Application");

  if(path === undefined) {
    this.workbookObj = this.excelObj.Workbooks.Add();
  } else {
    if(!File.exist(path)) {
      throw new Error(-1, "File Not Found: " + path);
    }

    switch(fileType) {
      case Excel.IN_FILETYPE_CSV:
      case Excel.IN_FILETYPE_TSV:
        var isFileTypeCsv = fileType === Excel.IN_FILETYPE_CSV;
        var isFileTypeTsv = fileType === Excel.IN_FILETYPE_TSV;

        var columnInfoSafeArray = Excel.array2dToSafeArray2d(columnInfo);
        this.excelObj.Workbooks.OpenText(path, 1, 1, 1, 1, false, isFileTypeTsv, false, isFileTypeCsv, false, false, false, columnInfoSafeArray);
        this.workbookObj = this.excelObj.ActiveWorkbook;
        break;
      default:
        this.workbookObj = this.excelObj.Workbooks.Open(path);
    }
  }
};

/**
 * ファイルフォーマット：xlCSV(6)
 */
Excel.FILE_FORMAT_CSV = 6;

/**
 * ファイルフォーマット：xlCurrentPlatformText(-4158)
 */
Excel.FILE_FORMAT_TSV = -4158;

/**
 * ファイルフォーマット：xlWorkbookNormal(-4143)
 */
Excel.FILE_FORMAT_EXCEL = -4143;

/**
 * 入力ファイルタイプ：タブ区切り(1)
 */
Excel.IN_FILETYPE_TSV = 1;

/**
 * 入力ファイルタイプ：セミコロン(2)
 */
Excel.IN_FILETYPE_SEMI_COLON = 2;

/**
 * 入力ファイルタイプ：CSV(4)
 */
Excel.IN_FILETYPE_CSV = 4;

/**
 * 入力ファイルタイプ：スペース(8)
 */
Excel.IN_FILETYPE_SPACE = 8;

/**
 * 入力ファイルタイプ：その他(16)
 */
Excel.IN_FILETYPE_OTHER = 16;

/**
 * カラムタイプ：xlGeneralFormat(1)
 */
Excel.COLUMN_TYPE_GENERAL = 1;

/**
 * カラムタイプ：xlTextFormat(2)
 */
Excel.COLUMN_TYPE_TEXT = 2;

/**
 * カラムタイプ：xlMDYFormat(3)
 */
Excel.COLUMN_TYPE_MDY = 3;

/**
 * カラムタイプ：xlDMYFormat(4)
 */
Excel.COLUMN_TYPE_DMY = 4;

/**
 * カラムタイプ：xlYMDFormat(5)
 */
Excel.COLUMN_TYPE_YMD = 5;

/**
 * カラムタイプ：xlMYDFormat(6)
 */
Excel.COLUMN_TYPE_MYD = 6;

/**
 * カラムタイプ：xlDYMFormat(7)
 */
Excel.COLUMN_TYPE_DYM = 7;

/**
 * カラムタイプ：xlYDMFormat(8)
 */
Excel.COLUMN_TYPE_YDM = 8;

/**
 * カラムタイプ：xlSkipColumn(9)
 */
Excel.COLUMN_TYPE_SKIP = 9;

/**
 * カラムタイプ：xlEMDFormat(10)
 */
Excel.COLUMN_TYPE_EMD = 10;

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
    if(excel) {
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
    if(excel) {
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
    return new Excel();
  }

  try {
    var excel = new Excel();
    block(excel);
  } finally {
    if(excel) {
      excel.quit();
    }
  }
};

/** 
 * Excelファイルを読み取り専用(最後に変更を破棄)で作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.createReadonly = function(block) {
  if(!isFunction(block)) {
    return new Excel();
  }

  try {
    var excel = new Excel();
    block(excel);
  } finally {
    if(excel) {
      excel.quitDiscardChanges();
    }
  }
};

/**
 * プログラムを動作させる端末でExcelが利用可能かどうか。
 * @return {Boolean} Excelが利用可能な場合はtrue、それ以外の場合はfalseを返す。
 */
Excel.available = function() {
  try {
    var x = new ActiveXObject("Excel.Application");
    return true;
  } catch(e) {
    return false;
  }
}

/**
 * ExcelブックをCSVファイルに変換する。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertToCsv = function(path) {
  Excel.openReadonly(path, function(excel) {
    excel.each(function(sheet) {
      sheet.activate();
      var outFileName = path + "_" + sheet.getName() + ".csv";
      excel.saveAsCsv(outFileName);
    });
  });
}

/**
 * Excelブックをタブ区切りファイルに変換する。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertToTsv = function(path) {
  Excel.openReadonly(path, function(excel) {
    excel.each(function(sheet) {
      sheet.activate();
      var outFileName = path + "_" + sheet.getName() + ".tsv";
      excel.saveAsTsv(outFileName);
    });
  });
}

/**
 * JScriptの二次元配列を、二次元のSafeArrayに変換する。
 * @param {Array} jsArray2d JScriptの二次元配列を指定します。
 * @return {Object} 変換された二次元のSafeArrayを返します。通常、VBArrayでラップして使います。<br/>
 *                  変換に失敗した場合はundefinedを返します。
 */
Excel.array2dToSafeArray2d = function(jsArray2d) {
  var safeArray = undefined;

  Excel.createReadonly(function(excel) {
    // 1シート目を取得
    var sheet = excel.getSheetByIndex(1);

    // 各セルに値を設定
    var i, j;
    for(i=0; i<jsArray2d.length; i++) {
      for(j=0; j<jsArray2d[i].length; j++) {
        var value = jsArray2d[i][j];
        sheet.setValue(i+1, j+1, value);
      }
    }

    // SafeArrayを取り出す。
    safeArray = sheet.getRawObject().Cells(1, 1).CurrentRegion.Value;
  });

  return safeArray;
}

/**
 * CSVファイルをExcelファイルに変換する。
 * @param {String} path CSVファイルのパスを文字列で指定します。
 * @param {Array} columnInfo (オプション)CSVファイルのカラム情報を指定します。<br>
<a href = "Excel.html#.COLUMN_TYPE_GENERAL">カラムタイプ</a>参照(初期値 全列に対しConst.COLUMN_TYPE_TEXT)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 * @example 使用例：
// 3カラム全て文字列として取り込む
var columnInfo = [
	[1, Excel.COLUMN_TYPE_TEXT]
	, [2, Excel.COLUMN_TYPE_TEXT]
	, [3, Excel.COLUMN_TYPE_TEXT]
];

// 拡張子がCSVだと、Excelが自動的に変換してうまくいかない
Excel.convertFromCsv("csv.txt", columnInfo);
 */
Excel.convertFromCsv = function(path, columnInfo) {
  var myPath = path;

  // 拡張子が.csvの場合は、うまく変換できないので、拡張子に.txtを付与したファイルにコピーする。
  if(path.endsWith(".csv")) {
    myPath = path + ".txt";
    File.copy(path, myPath);
  }

  var myColumnInfo = columnInfo;
  if(columnInfo === undefined) {
    myColumnInfo = Excel.generateDefaultColumnInfo(myPath, ",");
  }

  var excel = new Excel(myPath, Excel.IN_FILETYPE_CSV, myColumnInfo);
  excel.saveAsExcel(path + ".xls");
  excel.quit();

  // 一時ファイルを削除する
  if(path !== myPath) {
    File.unlink(myPath);
  }
}

/**
 * タブ区切りファイルをExcelファイルに変換する。
 * @param {String} path タブ区切りファイルのパスを文字列で指定します。
 * @param {Array} columnInfo (オプション)タブ区切りファイルのカラム情報を指定します。<br>
<a href = "Excel.html#.COLUMN_TYPE_GENERAL">カラムタイプ</a>参照(初期値 全列に対しConst.COLUMN_TYPE_TEXT)
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertFromTsv = function(path, columnInfo) {
  var myColumnInfo = columnInfo;
  if(columnInfo === undefined) {
    myColumnInfo = Excel.generateDefaultColumnInfo(path, "\t");
  }

  var excel = new Excel(path, Excel.IN_FILETYPE_TSV, myColumnInfo);
  excel.saveAsExcel(path + ".xls");
  excel.quit();
}

// 1行目の内容に基づいてデフォルトのcolumnInfoを生成する。
Excel.generateDefaultColumnInfo = function(path, separateChar) {
  var columnInfo = [];

  File.open(path, "r", function(file) {
    var firstLine = file.readLine();
    var columnCount = firstLine.split(separateChar).length;

    var i;
    for(i=0; i<columnCount; i++) {
      columnInfo.push([(i+1), Excel.COLUMN_TYPE_TEXT]);
    }
  });

  return columnInfo;
}

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
  getFullPathName: function(fileName) {
    var myFileName = fileName;
    if(myFileName.indexOf(File.PATH_SEPARATOR) === -1) {
      myFileName = File.realpath(myFileName);
    }

    return myFileName;
  },
  /** 
   * Excelファイルをファイル名を指定して保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAs: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName));
  },
  /** 
   * Excelファイルをファイル名を指定してCSV形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsCsv: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_CSV);
  },
  /** 
   * Excelファイルをファイル名を指定してタブ区切り形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsTsv: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_TSV);
  },
  /** 
   * Excelファイルをファイル名を指定してExcel形式で保存します。
   * @param {String} fileName 保存先のファイル名
   */
  saveAsExcel: function(fileName) {
    this.workbookObj.SaveAs(this.getFullPathName(fileName), Excel.FILE_FORMAT_EXCEL);
  },
  /**
   * Excelファイルを保存せずに閉じます。
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
  },
  /**
   * セル値をクリップボードへコピーします。
   * @param {Number} x x座標
   * @param {Number} y y座標
   */
  copy: function(x, y) {
    this.sheetObj.Cells(x, y).Copy();
  }
};
