function output(text, isOK) {
  var status = "OK";
  if(!isOK) {
    status = "NG";
  }
  WScript.Echo(status + ": " + text);
}

function assert(message, expr) {
  // argument processing
  var argumentsIndex = 0;
  if(arguments.length === 2) {
    var message = arguments[argumentsIndex++];
  } else {
    var message = "expected: true, actual: false";
  }
  var expr = arguments[argumentsIndex++];

  // assert
  if(!expr) {
    throw new Error(message);
  }

  assert.count++;
  return true;
}

function assertEquals() {
  // argument processing
  var argumentsIndex = 0;
  if(arguments.length === 3) {
    var message = arguments[argumentsIndex++];
  }
  var expected = arguments[argumentsIndex++];
  var actual = arguments[argumentsIndex++];
  if(!message) {
    message = "expected: " + expected + ", actual: " + actual;
  }

  // assert
  return assert(message, expected === actual);
}

function TestCase(name, tests) {
  assert.count = 0;
  var successful = 0;
  var testCount = 0;
  var hasSetup = typeof tests.setUp == "function";
  var hasTeardown = typeof tests.tearDown == "function";
  for(var test in tests) {
    if(!/^test/.test(test)) {
      continue;
    }

    testCount++;
    
    try {
      if(hasSetup) {
        tests.setUp();
      }
      tests[test]();
      output(test, true);
      if(hasTeardown) {
        tests.tearDown();
      }
      successful++;
    } catch(e) {
      output(test + " failed: " + e.message, false);
    }
  }

  output(testCount + " tests, " + (testCount - successful) + " failures", (testCount === successful));
}
