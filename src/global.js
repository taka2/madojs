// Global Functions
/**
 * 指定したmsec(ミリ秒)スリープします。
 * @param {Number} msec スリープするミリ秒
 */
var sleep = function(msec) {
  WScript.Sleep(msec);
};

var sleepif = function(msec, fn, args) {
  while(true) {
    if(fn(args)) {
      return true;
    }
    sleep(msec);
  }
};

var print = function(msg) {
  WScript.Echo(msg);
};

var exit = function(status) {
  if(!status) {
    status = 0;
  }

  WScript.Quit(status);
};

var alias = function (hash) {
  for (var prop in hash) {
    this[prop] = hash[prop];
  }
};

var extend = function (destination, source) {
  for (property in source) {
    destination[property] = source[property];
  }
  return destination;
};

var isFunction = function (fun) {
  return (typeof fun == "function");
};

var getComputerName = function () {
  return Const.WSHELL.ExpandEnvironmentStrings("%COMPUTERNAME%");
};

var isWScriptRunning = function() {
  return /wscript\.exe$/i.test(WScript.FullName);
};

var sendKeys = function (key) {
  Const.WSHELL.Sendkeys(key);
  return this;
};

var getEnv = function (name) {
  return Const.WSHELL.ExpandEnvironmentStrings("%" + name + "%");
};

var osShutdown = function (timeout /* タイムアウトをxx秒で指定 */) {
  if (timeout === null) timeout = 0;
  Process.exec("shutdown", ["-s", "-t", timeout], 0, false);
};

var osReboot = function (timeout /* タイムアウトをxx秒で指定 */, option /* -f(実行中のプロセスを警告なしで閉じる) */) {
  var params = ["-r", "-t"];
  if (timeout === null) timeout = 0;
  params.push(timeout);

  if (option !== null) {
    params.push(option);
  }

  Process.exec("shutdown", params, 0, false);
};
