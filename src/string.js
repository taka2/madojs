// Prototype of String
String.prototype.trim = function() {
  var text = this.valueOf();
  return (text || "").replace( /^\s+|\s+$/g, "" );
};
