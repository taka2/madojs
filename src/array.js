/** 
 * 各要素に対してブロックを評価します。
 * @param {Function} block ブロック
 * @example 使用例：
 * var x = [1, 2, 3].forEach(function(i) {
 *   return i*2;
 * });
 *
 * print(x); // 2,4,6
 */
Array.prototype.forEach = function(block) {
  var result = [];

  var thisLength = this.length;
  for(var i=0; i<thisLength; i++) {
    result.push(block(this[i]));
  }

  return result;
};

Array.prototype.each = Array.prototype.forEach;

/**
 * objが配列に含まれるかどうかチェックします。
 * @param {Object} obj 配列に含まれるかどうかチェックするオブジェクト
 * @return {Boolean} objが配列に含まれる場合true、そうでない場合falseを返す。
 */
Array.prototype.include = function(obj) {
  var thisLength = this.length;
  for(var i=0; i<thisLength; i++) {
    if(this[i] === obj) {
      return true;
    }
  }
  return false;
};
