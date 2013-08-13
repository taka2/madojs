/**
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったODBCデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoOdbcConnection.open("DSN1", "user", "password", function(odbccon) {
  var result = odbccon.executeQuery("SELECT * FROM TEST");

  result.each(function(row) {
    print(row["COL1"]); // 'X'
  });
});

</pre>
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoOdbcConnection = function(dsName, userName, password) {
  var connectString = "Provider=MSDASQL;DSN=" + dsName + ";";
  this.con = new AdoConnection(connectString, userName, password).getRawConnection();
};

/**
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} dsName データソース名
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoOdbcConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOdbcConnection(dsName, userName, password);
  }

  try {
    var con = new AdoOdbcConnection(dsName, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOdbcConnection.prototype = Object.create(AdoConnection.prototype);
AdoOdbcConnection.prototype.constructor = AdoOdbcConnection;
