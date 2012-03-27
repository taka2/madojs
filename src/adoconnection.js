/** 
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったデータベースへの接続を行うクラス。<br/>
          通常はこのクラスを直接使用せず、サブクラスを使います。
 * @param {String} connectString 接続文字列
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoConnection = function(connectString, userName, password) {
  this.con = new ActiveXObject("ADODB.Connection");
  this.con.Open(connectString, userName, password);
};

/**
 * Schema Enumの定数：Table(20)
 */
AdoConnection.AD_SCHEMA_TABLES = 20;

/**
 * DataTypeEnumの定数：adArray(0x2000)
 */
AdoConnection.AD_DATATYPE_ARRAY = 0x2000;
/**
 * DataTypeEnumの定数：adBigInt(20)
 */
AdoConnection.AD_DATATYPE_BIGINT = 20;
/**
 * DataTypeEnumの定数：adBinary(128)
 */
AdoConnection.AD_DATATYPE_BINARY = 128;
/**
 * DataTypeEnumの定数：adBoolean(11)
 */
AdoConnection.AD_DATATYPE_BOOLEAN = 11;
/**
 * DataTypeEnumの定数：adBSTR(8)
 */
AdoConnection.AD_DATATYPE_BSTR = 8;
/**
 * DataTypeEnumの定数：adChapter(136)
 */
AdoConnection.AD_DATATYPE_CHAPTER = 136;
/**
 * DataTypeEnumの定数：adChar(129)
 */
AdoConnection.AD_DATATYPE_CHAR = 129;
/**
 * DataTypeEnumの定数：adCurrency(6)
 */
AdoConnection.AD_DATATYPE_CURRENCY = 6;
/**
 * DataTypeEnumの定数：adDate(7)
 */
AdoConnection.AD_DATATYPE_DATE = 7;
/**
 * DataTypeEnumの定数：adDBDate(133)
 */
AdoConnection.AD_DATATYPE_DBDATE = 133;
/**
 * DataTypeEnumの定数：adDBTime(134)
 */
AdoConnection.AD_DATATYPE_DBTIME = 134;
/**
 * DataTypeEnumの定数：adDBTimeStamp(135)
 */
AdoConnection.AD_DATATYPE_DBTIMESTAMP = 135;
/**
 * DataTypeEnumの定数：adDecimal(14)
 */
AdoConnection.AD_DATATYPE_DECIMAL = 14;
/**
 * DataTypeEnumの定数：adDouble(5)
 */
AdoConnection.AD_DATATYPE_DOUBLE = 5;
/**
 * DataTypeEnumの定数：adEmpty(0)
 */
AdoConnection.AD_DATATYPE_EMPTY = 0;
/**
 * DataTypeEnumの定数：adError(10)
 */
AdoConnection.AD_DATATYPE_ERROR = 10;
/**
 * DataTypeEnumの定数：adFileTime(64)
 */
AdoConnection.AD_DATATYPE_FILETIME = 64;
/**
 * DataTypeEnumの定数：adGUID(72)
 */
AdoConnection.AD_DATATYPE_GUID = 72;
/**
 * DataTypeEnumの定数：adIDispatch(9)
 */
AdoConnection.AD_DATATYPE_IDISPATCH = 9;
/**
 * DataTypeEnumの定数：adInteger(3)
 */
AdoConnection.AD_DATATYPE_INTEGER = 3;
/**
 * DataTypeEnumの定数：adIUnknown(13)
 */
AdoConnection.AD_DATATYPE_IUNKNOWN = 13;
/**
 * DataTypeEnumの定数：adLongVarBinary(205)
 */
AdoConnection.AD_DATATYPE_LONGVARBINARY = 205;
/**
 * DataTypeEnumの定数：adLongVarChar(201)
 */
AdoConnection.AD_DATATYPE_LONGVARCHAR = 201;
/**
 * DataTypeEnumの定数：adLongVarWChar(203)
 */
AdoConnection.AD_DATATYPE_LONGVARWCHAR = 203;
/**
 * DataTypeEnumの定数：adNumeric(131)
 */
AdoConnection.AD_DATATYPE_NUMERIC = 131;
/**
 * DataTypeEnumの定数：adPropVariant(138)
 */
AdoConnection.AD_DATATYPE_PROPVARIANT = 138;
/**
 * DataTypeEnumの定数：adSingle(4)
 */
AdoConnection.AD_DATATYPE_SINGLE = 4;
/**
 * DataTypeEnumの定数：adSmallInt(2)
 */
AdoConnection.AD_DATATYPE_SMALLINT = 2;
/**
 * DataTypeEnumの定数：adTinyInt(16)
 */
AdoConnection.AD_DATATYPE_TINYINT = 16;
/**
 * DataTypeEnumの定数：adUnsignedBigInt(21)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDBIGINT = 21;
/**
 * DataTypeEnumの定数：adUnsignedInt(19)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDINT = 19;
/**
 * DataTypeEnumの定数：adUnsignedSmallInt(18)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDSMALLINT = 18;
/**
 * DataTypeEnumの定数：adUnsignedTinyInt(17)
 */
AdoConnection.AD_DATATYPE_UNSIGNEDTINYINT = 17;
/**
 * DataTypeEnumの定数：adUserDefined(132)
 */
AdoConnection.AD_DATATYPE_USERDEFINED = 132;
/**
 * DataTypeEnumの定数：adVarBinary(204)
 */
AdoConnection.AD_DATATYPE_VARBINARY = 204;
/**
 * DataTypeEnumの定数：adVarChar(200)
 */
AdoConnection.AD_DATATYPE_VARCHAR = 200;
/**
 * DataTypeEnumの定数：adVariant(12)
 */
AdoConnection.AD_DATATYPE_VARIANT = 12;
/**
 * DataTypeEnumの定数：adVarNumeric(139)
 */
AdoConnection.AD_DATATYPE_VARNUMERIC = 139;
/**
 * DataTypeEnumの定数：adVarWChar(202)
 */
AdoConnection.AD_DATATYPE_VARWCHAR = 202;
/**
 * DataTypeEnumの定数：adWChar(130)
 */
AdoConnection.AD_DATATYPE_WCHAR = 130;

/** 
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} connectString 接続文字列
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoConnection.open = function(connectString, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoConnection(connectString, userName, password);
  }

  try {
    var con = new AdoConnection(connectString, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoConnection
AdoConnection.prototype = {
  // private
  getRawConnection: function() {
    return this.con;
  },
  /**
   * コネクションの接続文字列を取得する。
   * @return {String} コネクションの接続文字列
   */
  getConnectionString: function() {
    return this.con.ConnectionString;
  },
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql) {
    try {
      // SQLの実行
      var rs = this.con.Execute(sql);
      return this.convertToArrayRs(rs);
    } catch(e) {
      throw e;
    } finally {
      if(rs) {
        rs.Close();
      }
    }
  },
  /**
   * 指定されたsqlを実行する。
   * @param {String} sql 実行するSQL
   */
  executeUpdate: function(sql) {
    // SQLの実行
    this.con.Execute(sql);
  },
  // private
  // Recordsetを配列にコンバートする。
  convertToArrayRs: function(rs) {
    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    while(!rs.EOF) {
      var record = {};
      for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
        var field = fe.item();
        record[field.Name] = field.Value;
      }
      result.push(record);

      rs.MoveNext();
    }

    return result;
  },
  /**
   * 指定されたsqlを実行するコマンドを作成します。
   * @param {String} sql 実行するSQL
   * @return {Object} コマンドオブジェクト
   */
  createCommand: function(sql) {
    return new AdoCommand(this, sql);
  },
  /**
   * データベースの接続を閉じる。
   */
  close: function() {
    this.con.Close();
  },
  /**
   * トランザクションを開始します。
   * @return {Number} トランザクションネストレベル
   */
  beginTrans: function() {
    return this.con.BeginTrans();
  },
  /**
   * トランザクションをコミットします。
   */
  commitTrans: function() {
    this.con.CommitTrans();
  },
  /**
   * トランザクションをロールバックします。
   */
  rollbackTrans: function() {
    this.con.RollbackTrans();
  },
  /**
   * テーブル名一覧を取得します。
   * @param {String} schema (オプション)スキーマを指定します。
   * @return {Array} テーブル名の配列を返します。
   */
  getTableNames: function(schema) {
    var rs = this.con.OpenSchema(AdoConnection.AD_SCHEMA_TABLES, arrayToSafeArray([undefined, schema, undefined, "TABLE"]));
    var array = this.convertToArrayRs(rs);
    var result = [];
    array.each(function(elem) {
      result.push(elem["TABLE_NAME"]);
    });

    return result;
  },
  /**
   * 指定したテーブルのフィールド名一覧を取得します。
   * @param {String} tableName フィールド名一覧を取得するテーブル名を指定します。
   * @return {Array} フィールド名一覧を返します。
   */
  getFieldNames: function(tableName) {
    // フィールドリスト取得のためのSQL空実行
    var rs = this.con.Execute("SELECT * FROM " + tableName);

    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
      result.push(fe.item().Name);
    }

    return result;
  },
  /**
   * 指定したテーブルのフィールド情報一覧を取得します。[{カラム名:カラムタイプ}]
   * @param {String} tableName フィールド情報一覧を取得するテーブル名を指定します。
   * @return {Array} フィールド情報一覧を返します。
   */
  getFieldInfo: function(tableName) {
    // フィールドリスト取得のためのSQL空実行
    var rs = this.con.Execute("SELECT * FROM " + tableName);

    // フィールドリストの取得
    var fe = new Enumerator(rs.Fields);

    // データの取得
    var result = [];

    for(fe.moveFirst(); !fe.atEnd(); fe.moveNext()) {
      var item = fe.item();
      result.push({'Name': item.Name, 'Type': item.Type});
    }

    return result;
  }
};
