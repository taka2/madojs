/** 
 * 日付をYYYYMMDD形式で返します。
 * @return {String} YYYYMMDD形式の日付
 * @example 使用例：
print(new Date().getYYYYMMDD());
 */
Date.prototype.getYYYYMMDD = function() {
  var yyyy = this.getFullYear();
  var mm = this.getMonth() + 1;
  var dd = this.getDate();
  if (mm < 10) mm = "0" + mm;
  if (dd < 10) dd = "0" + dd;
  return "" + yyyy + mm + dd;
};

/** 
 * 日付をYYYYMMDDHH24MISS形式で返します。
 * @return {String} YYYYMMDDHH24MISS形式の日付
 * @example 使用例：
print(new Date().getYYYYMMDDHH24MISS());
 */
Date.prototype.getYYYYMMDDHH24MISS = function() {
  var hh = this.getHours();
  var mi = this.getMinutes();
  var ss = this.getSeconds();
  if (hh < 10) hh = "0" + hh;
  if (mi < 10) mi = "0" + mi;
  if (ss < 10) ss = "0" + ss;
  return this.getYYYYMMDD() + hh + mi + ss;
};

/** 
 * 日付をYYYY/MM/DD HH:MI:SS形式で返します。
 * @return {String} YYYY/MM/DD HH:MM:SS形式の日付
 * @example 使用例：
print(new Date().getFormattedDate());
 */
Date.prototype.getFormattedDate = function() {
  var yyyy = this.getFullYear();
  var mm = this.getMonth() + 1;
  var dd = this.getDate();
  var hh = this.getHours();
  var mi = this.getMinutes();
  var ss = this.getSeconds();
  if (mm < 10) mm = "0" + mm;
  if (dd < 10) dd = "0" + dd;
  if (hh < 10) hh = "0" + hh;
  if (mi < 10) mi = "0" + mi;
  if (ss < 10) ss = "0" + ss;
  return yyyy + "/" + mm + "/" + dd + " " + hh + ":" + mi + ":" + ss;
};
