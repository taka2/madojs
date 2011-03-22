// �e�X�g�N���X���i�[����O���[�o���I�u�W�F�N�g
var Test = {};

// �e�X�g���[�e�B���e�B�N���X
var TestUtil = function() {
  this.totalTestsCount = 0;
  this.totalOkCount = 0;
  this.totalNgCount = 0;
};

// �e�X�g�p���\�b�h
TestUtil.prototype.assertEquals = function(a, b) {
  this.totalTestsCount++;

  try {
    if(a === b) {
      this.totalOkCount++;
    } else {
      this.totalNgCount++;
    }
  } catch (e) {
    this.totalNgCount++;
  }
};

// �e�X�g�̎��s
TestUtil.prototype.runTest = function() {
  for(var cls in Test) {
    for(var testMethod in Test[cls]) {
      var objTestMethod = Test[cls][testMethod];
      if(typeof(objTestMethod) === "function") {
        // �֐��̏ꍇ�̂݃R�[������
        objTestMethod();
      }
    }
  }
};

// �e�X�g���ʂ̏o��
TestUtil.prototype.printTestResult = function() {
  print(this.totalOkCount + "/" + this.totalTestsCount + " is Ok.");
};