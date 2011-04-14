/** 
 * 新しいADOのコマンドを作成する。<br/>
   通常、AdoCommandを直接インスタンス化せず、AdoConnectionのサブクラスから、createCommandを呼びます。
 * @class ADOコマンドを表すクラス。クエリにパラメータが必要な場合は、このクラスを使います。
<pre class = "code">
使用例：
AdoAccessConnection.open("c:\\path\\to\\test.mdb", "", "", function(con) {
  var command = con.createCommand("UPDATE TEST SET COL2 = ? WHERE COL1 = ?");
  command.setStringParameter("COL2", "えええ");
  command.setStringParameter("COL1", "Z");
  command.executeUpdate();
});

AdoAccessConnection.open("c:\\path\\to\\test.mdb", "", "", function(con) {
  var command = con.createCommand("SELECT COL2 FROM TEST WHERE COL1 = ?");
  command.setStringParameter("COL1", "Z");
  var rs = command.executeQuery();

  print(rs[0]["COL2"]);
});

</pre>
 * @param {Object} con Adoコネクション
 */
var AdoCommand = function(con, sql) {
  this.con = con;
  this.sql = sql;
  this.params = [];

  // コマンドの生成
  this.command = new ActiveXObject("ADODB.Command");
  this.command.CommandText = this.sql;
  this.command.CommandType = AdoCommand.COMMAND_TYPE_CMD_TEXT;
  this.command.ActiveConnection = this.con.getRawConnection();
};

/**
 * Command Typeの定数：テキスト(1)
 */
AdoCommand.COMMAND_TYPE_CMD_TEXT = 1; // 通常

/**
 * データ型の定数: Boolean(11)
 */
AdoCommand.DATA_TYPE_BOOLEAN = 11;

/**
 * データ型の定数: Char(129)
 */
AdoCommand.DATA_TYPE_CHAR = 129;

/**
 * データ型の定数: DBDate(133)
 */
AdoCommand.DATA_TYPE_DBDATE = 133;

/**
 * データ型の定数: DBTime(134)
 */
AdoCommand.DATA_TYPE_DBTIME = 134;

/**
 * データ型の定数: DBTimestamp(135)
 */
AdoCommand.DATA_TYPE_DBTIMESTAMP = 135;

/**
 * データ型の定数: Decimal(14)
 */
AdoCommand.DATA_TYPE_DECIMAL = 14;

/**
 * データ型の定数: Double(5)
 */
AdoCommand.DATA_TYPE_DOUBLE = 5;

/**
 * データ型の定数: Integer(3)
 */
AdoCommand.DATA_TYPE_INTEGER = 3;

/**
 * データ型の定数: Numeric(131)
 */
AdoCommand.DATA_TYPE_NUMERIC = 131;

/**
 * データ型の定数: Single(4)
 */
AdoCommand.DATA_TYPE_SINGLE = 4;

/**
 * パラメータ方向の定数: Input(1)
 */
AdoCommand.PARAMETER_DIRECTION_INPUT = 1;

// Prototypes of AdoCommand
AdoCommand.prototype = {
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @param {Array} params パラメータ
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql, params) {
    try {
      // パラメータの追加
      this.prepareParameter();

      // SQLの実行
      var rs = this.command.Execute();
      return this.con.convertToArrayRs(rs);
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
   * @param {Array} params パラメータ
   */
  executeUpdate: function(sql, params) {
    // パラメータの追加
    this.prepareParameter();

    // SQLの実行
    this.command.Execute();
  },
  // private
  // パラメータの準備をする。
  prepareParameter: function() {
    // パラメータの追加
    var command = this.command;
    this.params.each(function(param) {
      var objParam = command.CreateParameter(
        param.Name
        , param.Type
        , param.Direction
        , param.Size
        , param.Value
      );
      command.Parameters.Append(objParam);
    });
  },
  /**
   * 論理型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Boolean} paramValue パラメータ値
   */
  setBooleanParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_BOOLEAN
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 文字列型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {String} paramValue パラメータ値
   */
  setStringParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_CHAR
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 日付型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setDateParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBDATE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * 時刻型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setTimeParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBTIME
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Timestamp型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setTimestampParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DBTIMESTAMP
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Deciaml型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Number} paramValue パラメータ値
   */
  setDecimalParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DECIMAL
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Double型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setDoubleParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_DOUBLE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Integer型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setIntegerParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_INTEGER
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Numeric型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setNumericParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_NUMERIC
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  },
  /**
   * Single型のパラメータを作成する。
   * @param {String} paramName パラメータ名
   * @param {Date} paramValue パラメータ値
   */
  setSingleParameter: function(paramName, paramValue) {
    this.params.push({
      "Name": paramName
      , "Type": AdoCommand.DATA_TYPE_SINGLE
      , "Direction": AdoCommand.PARAMETER_DIRECTION_INPUT
      , "Size": -1
      , "Value": paramValue
    });
  }
};
