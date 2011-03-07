/** 
 * 新しいADOのOracleコネクションを作成する
 * @class ADOを使ったOracleデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoOracleConnection.open("ORCL", "scott", "tiger", function(oracon) {
  var result = oracon.executeQuery("SELECT * FROM DUAL");

  result.each(function(row) {
    print(row["DUMMY"]); // 'X'
  });
});

</pre>
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
AdoOracleConnection = function(dsName, userName, password) {
  this.con = new ActiveXObject("ADODB.Connection");

  // Open Connection
  this.con.Open("Provider=MSDAORA;Data Source=" + dsName + ";", userName, password);
};

/** 
 * 新しいADOのOracleコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのOracleコネクションを返します。
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのOracleコネクション
 */
AdoOracleConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOracleConnection(dsName, userName, password);
  }

  try {
    var oracon = new AdoOracleConnection(dsName, userName, password);
    block(oracon);
  } finally {
    if (oracon != null) {
      oracon.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOracleConnection.prototype = {
  /**
   * 指定されたsqlを実行し、ハッシュ{fieldName: value}の配列として値を返す。
   * @param {String} sql 実行するSQL
   * @return {Array} ハッシュ{fieldName: value}の配列
   */
  executeQuery: function(sql) {
    try {
      // SQLの実行
      var rs = this.con.Execute(sql);

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
    } catch(e) {
      throw e;
    } finally {
      if(rs) {
        rs.Close();
      }
    }
  },
  close: function() {
    this.con.Close();
  }
};
