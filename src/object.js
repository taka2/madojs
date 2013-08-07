/** 
 * オブジェクトの値を表す文字列を返します。
 * @return {String} オブジェクトの値を表す文字列
 */
Object.prototype.toString = function() {
  var result = "{";
  for(e in this) {
    if(this.hasOwnProperty(e)) {
      if(result !== "{") {
        result += ", ";
      }
      result += '"' + e + '": ' + this[e].toString();
    }
  }
  result += "}";
  return result;
};

/**
 * pをプロトタイプに持つオブジェクトを生成します。
 * @param {Object} p プロトタイプ
 * @return {Object} pをプロトタイプに持つオブジェクト
 */
Object.prototype.create = function(p) {
  function f() {};
  f.prototype = p;
  return new f();
};
