// テストクラスを格納するグローバルオブジェクト
var Test = {};

// テストユーティリティクラス
var TestUtil = function() {
  this.totalTestsCount = 0;
  this.totalOkCount = 0;
  this.totalNgCount = 0;
  this.errors = [];
  this.currentMethodName = "";
};

// テスト用メソッド
TestUtil.prototype.assertEquals = function(a, b) {
  this.totalTestsCount++;

  try {
    if(a === b) {
      this.totalOkCount++;
    } else {
      this.totalNgCount++;
      this.errors.push(this.currentMethodName + ": expected: " + a + ", actual: " + b);
    }
  } catch (e) {
    this.totalNgCount++;
  }
};

// テストの実行
TestUtil.prototype.runTest = function() {
  for(var cls in Test) {
    for(var testMethod in Test[cls]) {
      var objTestMethod = Test[cls][testMethod];
      if(typeof(objTestMethod) === "function") {
        // 関数の場合のみコールする
        this.currentMethodName = "Test." + cls + "." + testMethod;
        objTestMethod();
      }
    }
  }
};

// テスト結果の出力
TestUtil.prototype.printTestResult = function() {
  print(this.totalOkCount + "/" + this.totalTestsCount + " is Ok.");

  var errorMessage = "";
  for(var i=0; i<this.errors.length; i++) {
    errorMessage = errorMessage + this.errors[i] + "\n";
  }

  if(errorMessage !== "") {
    print(errorMessage);
  }
};
