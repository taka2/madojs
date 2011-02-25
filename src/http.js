/** 
 * hostとportを指定してHTTPオブジェクトを作成します。
 * @class HTTPの操作を行うためのクラスです。
<pre class = "code">
使用例：
HTTP.start("www.yahoo.co.jp", 80, function(http) {
  print(http.get("/index.html"));
});
</pre>
 * @param {String} host HTTP操作の対象ホスト
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 */
var HTTP = function(host, port) {
  this.host = host;
  this.port = port || DEFAULT_PORT_NUMBER;
};

/**
 * 定数：デフォルトのポート番号(80)
 */
HTTP.DEFAULT_PORT_NUMBER = 80;

/** 
 * hostとportを指定してHTTPオブジェクトを作成し、ブロックを実行します。
 * ブロックが指定されていない場合は、作成したHTTPオブジェクトを返します。
 * @param {String} host HTTP操作の対象ホスト
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @param {Function} block (オプション)ブロック
 * @return {Object} ブロックが指定されていない場合は、作成したHTTPオブジェクト
 */
HTTP.start = function(host, port, block) {
  var http = new HTTP(host, port);
  if(block) {
    block(http);
  } else {
    return http;
  }
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行います。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @return {String} レスポンステキスト
 * @example 使用例：
print(HTTP.get("www.yahoo.co.jp", "/index.html", 80));
 */
HTTP.get = function(address, path, port) {
  var myPort = port || DEFAULT_PORT_NUMBER;
  var http = new HTTP(address, port);
  return http.get(path);
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行い、printします。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @example 使用例：
HTTP.get_print("www.yahoo.co.jp", "/index.html", 80);
 */
HTTP.get_print = function(address, path, port) {
  print(HTTP.get(address, path, port));
};

// Prototypes of HTTP
HTTP.prototype = {
  /** 
   * 指定されたpath、headerでGETリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  get: function(path, header, block) {
    // パラメータの調整
    if(header === undefined) {
      header = [];
    }

    // リクエストの準備
    var xhr = new ActiveXObject("Microsoft.XMLHTTP");
    xhr.open("GET", "http://" + this.host + ":" + this.port + path, false);

    for(var headerName in header) {
      xhr.setRequestHeader(headerName, header[headerName]);
    }

    // リクエストの送信
    xhr.send();

    // レスポンスの処理
    if(xhr.readyState == 4) {
      if(xhr.status == 200) {
        if(block) {
          block(xhr.responseText);
        } else {
          return xhr.responseText;
        }
      } else {
        throw new Error(xhr.statusText);
      }
    }
  }
};
