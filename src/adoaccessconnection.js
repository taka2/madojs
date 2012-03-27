/** 
 * 新しいADOのコネクションを作成する
 * @class ADOを使ったAccessデータベースへの接続を行うクラス
<pre class = "code">
使用例：
AdoAccessConnection.open("c:\\path\\to\\test.mdb", "hoge", "fuga", function(con) {
  var result = con.executeQuery("SELECT * FROM TEST");

  result.each(function(row) {
    print(row["COL1"]); // 'X'
  });
});

</pre>
 * @param {String} mdbFilePath mdbファイルのパス
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 */
var AdoAccessConnection = function(mdbFilePath, userName, password) {
  var connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbFilePath;
  this.con = new AdoConnection(connectString, userName, password).getRawConnection();
};

/** 
 * 新しいADOのコネクションを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したADOのコネクションを返します。
 * @param {String} mdbFilePath mdbファイルのパス
 * @param {String} userName ユーザ名
 * @param {String} password パスワード
 * @param {Function} block ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したADOのコネクション
 */
AdoAccessConnection.open = function(mdbFilePath, userName, password, block) {
  if (!isFunction(block)) {
    return new AdoAccessConnection(mdbFilePath, userName, password);
  }

  try {
    var con = new AdoAccessConnection(mdbFilePath, userName, password);
    block(con);
  } finally {
    if(con) {
      con.close();
    }
  }
};

// Prototypes of AdoAccessConnection
AdoAccessConnection.prototype = AdoConnection.prototype;
