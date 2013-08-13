// Collections
// each
_.each([1, 2, 3], print);

// map
print(_.map([1, 2, 3], function(num) { return num * 3; }));

// reduce
print(_.reduce([1, 2, 3], function(memo, num){ return memo + num; }, 0));

// Arrays
// first
print(_.first([5, 4, 3, 2, 1]));

// initial
print(_.initial([5, 4, 3, 2, 1]));

// last
print(_.last([5, 4, 3, 2, 1]));

// Functions
// bind
var func = function(greeting){ return greeting + ":" + this.name };
func = _.bind(func, {name : 'moe'}, 'hi');
print(func());
