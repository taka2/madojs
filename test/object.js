Test.Object = {};

Test.Object.testToString = function() {
  var obj1 = {"a":"str"};
  testUtil.assertEquals('{"a": "str"}', obj1.toString());

  var obj2 = {"a":123};
  testUtil.assertEquals('{"a": 123}', obj2.toString());

  var obj3 = {"a":1.23};
  testUtil.assertEquals('{"a": 1.23}', obj3.toString());

  var obj4 = {"a":true};
  testUtil.assertEquals('{"a": true}', obj4.toString());

  var obj5 = {'a':[1,2,'ccc']};
  testUtil.assertEquals('{"a": [1, 2, "ccc"]}', obj5.toString());

  var obj7 = {"a":{"b": 1.23}};
  testUtil.assertEquals('{"a": {"b": 1.23}}', obj7.toString());

  var obj8 = {"a": {"b": {"c": true}}};
  testUtil.assertEquals('{"a": {"b": {"c": true}}}', obj8.toString());
};
