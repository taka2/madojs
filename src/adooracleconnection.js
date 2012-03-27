/**
 * 新しいADOのコネクションを作成する
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
var AdoOracleConnection = function(dsName, userName, password) {
  var connectString = "Provider=MSDAORA;Data Source=" + dsName + ";";
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
AdoOracleConnection.open = function(dsName, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoOracleConnection(dsName, userName, password);
  }

  try {
    var con = new AdoOracleConnection(dsName, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoOracleConnection
AdoOracleConnection.prototype = AdoConnection.prototype;
