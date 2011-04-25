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
  }
};
