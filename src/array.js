/** 
 * 各要素に対してブロックを評価します。
 * @param {Function} block ブロック
 * @example 使用例：
 * var x = [1, 2, 3].each(function(i) {
 *   return i*2;
 * });
 *
 * print(x); // 2,4,6
 */
Array.prototype.each = function(block) {
  var result = [];

  for(var i=0; i<this.length; i++) {
    result.push(block(this[i]));
  }

  return result;
};
