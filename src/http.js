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
 * 定数：レスポンス形式（テキスト）
 */
HTTP.RESPONSE_TYPE_TEXT = "TEXT";
/**
 * 定数：レスポンス形式（ボディ）
 */
HTTP.RESPONSE_TYPE_BODY = "BODY";
/**
 * 定数：レスポンス形式（XML）
 */
HTTP.RESPONSE_TYPE_XML = "XML";
/**
 * 定数：XHR実装クラス：Microsoft.XMLHTTP
 */
HTTP.XHR_IMPL_CLASS_Microsoft_XMLHTTP = "Microsoft.XMLHTTP";
/**
 * 定数：XHR実装クラス：MSXML2.XMLHTTP5.0
 */
HTTP.XHR_IMPL_CLASS_MSXML2_XMLHTTP5_0 = "MSXML2.XMLHTTP5.0";

/**
 * グローバル変数：XHR実装クラス(初期値 XHR_IMPL_CLASS_Microsoft_XMLHTTP
 */
HTTP.XHR_IMPL_CLASS = HTTP.XHR_IMPL_CLASS_Microsoft_XMLHTTP;

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
 * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
 * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
 * @return {String} レスポンス
 * @example 使用例：
print(HTTP.get("www.yahoo.co.jp", "/index.html", 80));
 */
HTTP.get = function(address, path, port, header, responseType) {
  var myPort = port || DEFAULT_PORT_NUMBER;
  var myHeader = header || {};
  var myResponseType = responseType || HTTP.RESPONSE_TYPE_TEXT;
  var http = new HTTP(address, myPort);
  return http.get(path, myHeader, myResponseType);
};

/** 
 * 指定されたaddress、path、portに対して、GETリクエストを行い、printします。
 * @param {String} address HTTP操作の対象ホスト
 * @param {String} path HTTP操作の対象パス
 * @param {Number} port (オプション)HTTP操作の対象ポート(初期値 80)
 * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
 * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
 * @example 使用例：
HTTP.get_print("www.yahoo.co.jp", "/index.html", 80);
 */
HTTP.get_print = function(address, path, port, header, responseType) {
  print(HTTP.get(address, path, port, header, responseType));
};

// Prototypes of HTTP
HTTP.prototype = {
  /** 
   * 指定されたpath、headerでGETリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ(初期値 {})
   * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  get: function(path, header, responseType, block) {
    return this.request("GET", path, header, responseType, "", block);
  },
  /** 
   * 指定されたpath、headerでPOSTリクエストを行い、ブロックを実行します。
   * ブロックが指定されていない場合は、レスポンステキストを返します。
   * @param {String} path HTTP操作の対象パス
   * @param {Object} header (オプション)HTTPヘッダ(初期値 [])
   * @param {String} responseType (オプション)レスポンス形式(初期値 TEXT)
   * @param {Object} body (オプション)HTTPボディ(初期値 undefined)
   * @param {Function} block (オプション)ブロック
   * @return {String} ブロックが指定されていない場合は、レスポンステキスト
   * @throws 200OK以外のレスポンスが返された場合にスローされます。
   */
  post: function(path, header, responseType, body, block) {
    return this.request("POST", path, header, responseType, body, block);
  },
  // private
  request: function(method, path, header, responseType, body, block) {
    // パラメータの調整
    var myHeader = header || {};

    // リクエストの準備
    var xhr = new ActiveXObject(HTTP.XHR_IMPL_CLASS);
    xhr.open(method, "http://" + this.host + ":" + this.port + path, false);

    for(var headerName in myHeader) {
      if(myHeader.hasOwnProperty(headerName)) {
        xhr.setRequestHeader(headerName, myHeader[headerName]);
      }
    }

    // リクエストの送信
    xhr.send(body);

    // レスポンスの処理
    if(xhr.readyState === 4) {
      if(xhr.status === 200) {
        var response;
        if(responseType === HTTP.RESPONSE_TYPE_TEXT) {
          response = xhr.responseText;
        } else if(responseType === HTTP.RESPONSE_TYPE_BODY) {
          response = xhr.responseBody;
        } else if(responseType === HTTP.RESPONSE_TYPE_XML) {
          response = xhr.responseXML;
        } else {
          response = xhr.responseText;
        }
        if(isFunction(block)) {
          block(response);
        } else {
          return response;
        }
      } else {
        throw new Error(xhr.statusText);
      }
    }
  }
};
