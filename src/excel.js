/** 
 * Excelオブジェクトを作成します。
 * @class Excelファイルの操作と読み書きを行うためのクラスです。<br/>
 * Excelがインストールされている必要があります。
<pre class = "code">
使用例：
// abc.xlsを開いて、Sheet1の、1行1列目の値を'hoge'に書き換えつつ、上書き保存する。
Excel.open("abc.xls", {}, function(excel) {
  var sheet = excel.getSheetByName("Sheet1");
  sheet.setValue(1, 1, 'hoge');
  excel.save();
});
</pre>
 * @param {String} path Excelファイルのパスを文字列で指定します。指定しない場合は新しくブックを作成します。
 * @param {Object} options (オプション)Excelファイルを開くときのオプションを指定します。
 * @param {Boolean} fileType (オプション)Excel以外のファイルを開くときに、指定します。
 * @param {Boolean} columnInfo (オプション)Excel以外のファイルを開くときに、カラムの情報を指定します。
 * @throws pathがundefinedでない、かつ、ファイルが存在しない場合にスローされます。
 */
var Excel = function(path, options, fileType, columnInfo) {
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
      case Excel.IN_FILETYPE_FIXED:
        var isFileTypeCsv = fileType === Excel.IN_FILETYPE_CSV;
        var isFileTypeTsv = fileType === Excel.IN_FILETYPE_TSV;
        var textParsingType = fileType === Excel.IN_FILETYPE_FIXED ? Excel.TEXT_PARSING_TYPE_FIXED_WIDTH : Excel.TEXT_PARSING_TYPE_DELIMITED;

        var columnInfoSafeArray = Excel.array2dToSafeArray2d(columnInfo);
        this.excelObj.Workbooks.OpenText(path, 932, 1, textParsingType, 1, false, isFileTypeTsv, false, isFileTypeCsv, false, false, false, columnInfoSafeArray);
        this.workbookObj = this.excelObj.ActiveWorkbook;
        break;
      default:
        var updateLinks = options[Excel.WORKBOOKS_OPEN_OPTION_UPDATE_LINKS] || false;
        var readOnly = options[Excel.WORKBOOKS_OPEN_OPTION_READ_ONLY] || false;
        this.workbookObj = this.excelObj.Workbooks.Open(path, updateLinks, readOnly);
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
 * 入力ファイルタイプ：固定長(32)
 */
Excel.IN_FILETYPE_FIXED = 32;

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
 * 印刷の向き：xlLandscape(2)
 */
Excel.PAGE_ORIENTATION_LANDSCAPE = 2;

/**
 * 印刷の向き：xlPortrait(1)
 */
Excel.PAGE_ORIENTATION_PORTRAIT = 1;

/**
 * Workbooks.Openオプション：UpdateLinks
 */
Excel.WORKBOOKS_OPEN_OPTION_UPDATE_LINKS = "UpdateLinks";

/**
 * Workbooks.Openオプション：ReadOnly
 */
Excel.WORKBOOKS_OPEN_OPTION_READ_ONLY = "ReadOnly";

/**
 * テキストパースタイプ：xlDelimited(1)
 */
Excel.TEXT_PARSING_TYPE_DELIMITED = 1;

/**
 * テキストパースタイプ：xlFixedWidth(2)
 */
Excel.TEXT_PARSING_TYPE_FIXED_WIDTH = 2;

/** 
 * Excelファイルを開き、ブロックを実行します。
 * ブロックが指定されていない場合は、Excelオブジェクトを返します。
 * @param {String} path Excelファイルのパスを文字列で指定します。
 * @param {Object} options (オプション)Excelファイルを開くときのオプションを指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.open = function(path, options, block) {
  if(!isFunction(block)) {
    return new Excel(path, options);
  }

  try {
    var excel = new Excel(path, options);
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
 * @param {Object} options (オプション)Excelファイルを開くときのオプションを指定します。
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したクリップボード
 */
Excel.openReadonly = function(path, options, block) {
  var myOptions = options || {};
  myOptions[Excel.WORKBOOKS_OPEN_OPTION_READ_ONLY] = true;

  if(!isFunction(block)) {
    return new Excel(path, myOptions);
  }

  try {
    var excel = new Excel(path, myOptions);
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
  Excel.openReadonly(path, undefined, function(excel) {
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
  Excel.openReadonly(path, undefined, function(excel) {
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
    var sheet = excel.getSheetByIndex(0);

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

  var excel = new Excel(myPath, undefined, Excel.IN_FILETYPE_CSV, myColumnInfo);
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

  var excel = new Excel(path, undefined, Excel.IN_FILETYPE_TSV, myColumnInfo);
  excel.saveAsExcel(path + ".xls");
  excel.quit();
}

/**
 * 固定長ファイルをExcelファイルに変換する。
 * @param {String} path 固定長ファイルのパスを文字列で指定します。
 * @param {Array} columnInfo 固定長ファイルのカラム情報を指定します。
 * @return {Boolean} 成功した場合はtrue、失敗した場合はfalseを返す。
 */
Excel.convertFromFixed = function(path, columnInfo) {
  var excel = new Excel(path, undefined, Excel.IN_FILETYPE_FIXED, columnInfo);
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
   * @param {Number} index 取得するワークシートのインデックス番号(0オリジン)
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
  },
  /**
   * 先頭から指定した枚数のシートを削除します。
   * @param {Number} numSheet 削除するシート数
   */
  removeSheets: function(numSheets) {
    var i;
    for(i=0; i<numSheets; i++) {
      var sheet = this.getSheetByIndex(0);
      sheet.remove();
    }
  },
  /**
   * 先頭のシートを選択します。
   */
  activateHeadSheet: function() {
    this.getSheetByIndex(0).activate();
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
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Object} value 設定する値
   */
  setValue: function(row, col, value) {
    this.sheetObj.Cells(row, col).Value = value;
  },
  /**
   * セルに値を設定します。
   * @param {Number} row 行
   * @param {Array} value 設定する値
   * @param {Number} coloffset (オプション)列の開始位置(初期値 先頭列)
   */
  setArrayValue: function(row, value, coloffset) {
    if(!coloffset) {
      coloffset = 1;
    }
    var valueLength = value.length;
    for(var i=0; i<valueLength; i++) {
      this.sheetObj.Cells(row, i+coloffset).Value = value[i];
    }
  },
  /**
   * セルの値を取得します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @return {Object} セルの値
   */
  getValue: function(row, col) {
    return this.sheetObj.Cells(row, col).Value;
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
   * Cell(1, 1)のカレントリージョンに対して罫線を引きます。
   */
  drawBorder: function() {
    var range = this.sheetObj.Cells(1, 1).CurrentRegion;
    for(var i=7; i<=10; i++) {
      range.Borders(i).LineStyle = 1;
    }
    if(range.Columns.Count > 1) {
        range.Borders(11).LineStyle = 1;
    }
    if(range.Rows.Count > 1) {
        range.Borders(12).LineStyle = 1;
    }
  },
  /**
   * セルの書式を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} format 書式
   */
  setFormat: function(row, col, format) {
    this.sheetObj.Cells(row, col).NumberFormatLocal = format;
  },
  /**
   * セルにコメントを追加します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} comment コメント
   * @param {Boolean} visible コメントを表示するかどうか
   */
  addComment: function(row, col, comment, visible) {
    this.sheetObj.Cells(row, col).AddComment();
    this.sheetObj.Cells(row, col).Comment.Visible = visible;
    this.sheetObj.Cells(row, col).Comment.Text(comment);
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
   * @param {Number} row 行
   * @param {Number} col 列
   */
  copy: function(row, col) {
    this.sheetObj.Cells(row, col).Copy();
  },
  /**
   * セルの背景色を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} colorIndex カラーインデックス
   */
  setBackgroundColor: function(row, col, colorIndex) {
    this.sheetObj.Cells(row, col).Interior.ColorIndex = colorIndex;
  },
  /**
   * セルのフォント色を設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} colorIndex カラーインデックス
   */
  setFontColor: function(row, col, colorIndex) {
    this.sheetObj.Cells(row, col).Font.ColorIndex = colorIndex;
  },
  /**
   * セルのフォントを太字に設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isBold 太字にするかどうか
   */
  setFontBold: function(row, col, isBold) {
    this.sheetObj.Cells(row, col).Font.Bold = isBold;
  },
  /**
   * セルのフォントをイタリックに設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isItalic 太字にするかどうか
   */
  setFontItalic: function(row, col, isItalic) {
    this.sheetObj.Cells(row, col).Font.Italic = isItalic;
  },
  /**
   * セルのフォントに取消し線を設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isStrikethrough 取消し線を設定するかどうか
   */
  setFontStrikethrough: function(row, col, isStrikethrough) {
    this.sheetObj.Cells(row, col).Font.Strikethrough = isStrikethrough;
  },
  /**
   * セルのフォントに下線を設定または解除します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {Number} isUnderline 下線を設定するかどうか
   */
  setFontUnderline: function(row, col, isUnderline) {
    this.sheetObj.Cells(row, col).Font.Underline = isUnderline;
  },
  /**
   * セルを選択します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  selectCell: function(row, col) {
    this.sheetObj.Cells(row, col).Select();
  },
  /**
   * セル範囲を選択します。
   * @param {Number} row1 行（始点）
   * @param {Number} col1 列（始点）
   * @param {Number} row2 行（終点）
   * @param {Number} col2 列（終点）
   */
  selectRange: function(row1, col1, row2, col2) {
    this.sheetObj.Range(this.sheetObj.Cells(row1, col1), this.sheetObj.Cells(row2, col2)).Select();
  },
  /**
   * セルにハイパーリンクを設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   * @param {String} address リンク先
   * @param {String} displayText (オプション)表示文字列(初期値 リンク先)
   */
  addHyperLink: function(row, col, address, displayText) {
    if(!displayText) {
      displayText = address;
    }
    var anchor = this.sheetObj.Cells(row, col);
    this.sheetObj.Hyperlinks.Add(anchor, address, "", address, displayText);
  },
  /**
   * セルに設定されたハイパーリンクを削除します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  deleteHyperLink: function(row, col) {
    this.sheetObj.Cells(row, col).Hyperlinks.Delete();
  },
  /**
   * 指定した列のカラム幅を設定します。
   * @param {Number} col 列(1オリジン)
   * @param {Number} columnWidth 列幅
   */
  setColumnWidth: function(col, columnWidth) {
    this.sheetObj.Columns(col).ColumnWidth = columnWidth;
  },
  /**
   * 指定したセルにオートフィルタを設定します。
   * @param {Number} row 行
   * @param {Number} col 列
   */
  setAutoFilter: function(row, col) {
    this.sheetObj.Cells(row, col).CurrentRegion.AutoFilter();
  },
  /**
   * 印刷の向きを設定します。
   * @param {Number} orientation 印刷の向き
   */
  setPageOrientation: function(orientation) {
    this.sheetObj.PageSetup.Orientation = orientation;
  },
  /**
   * ページ数に合わせて印刷を設定します。
   * @param {Number} wide 横のページ数、-1を指定した場合は指定なし。
   * @param {Number} tall 縦のページ数、-1を指定した場合は指定なし。
   */
  setPageFitToPages: function(wide, tall) {
    var myWide = wide == -1 ? false : wide;
    var myTall = tall == -1 ? false : tall;

    this.sheetObj.PageSetup.Zoom = false;
    this.sheetObj.PageSetup.FitToPagesWide = myWide;
    this.sheetObj.PageSetup.FitToPagesTall = myTall;
  },
  /**
   * 印刷の拡大率を設定します。
   * @param {Number} zoom 印刷の拡大率
   */
  setPageZoom: function(zoom) {
    this.sheetObj.PageSetup.Zoom = zoom;
  }
};
