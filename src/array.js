/** 
 * 各要素に対してブロックを評価します。
 * @param {Function} block ブロック
 */
Array.prototype.each = function(block) {
  var result = [];

  for(var i=0; i<this.length; i++) {
    result.push(block(this[i]));
  }

  return result;
};
