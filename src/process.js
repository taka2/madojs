// Constructor of Process
var Process = {};

// Static methods of File
Process.exec = function(command, args, windowStyle, waitOnReturn) {
  if(!args) {
    args = [];
  }
  if(!windowStyle) {
    windowStyle = Const.WINDOW_STYLE_NORMAL;
  }
  if(waitOnReturn === undefined) {
    waitOnReturn = true;
  }
  var commandLine = command + " " + args.join(" ");
  Const.WSHELL.Run(commandLine, windowStyle, waitOnReturn);
};
