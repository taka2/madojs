@echo off

REM src配下のファイルをまとめる
REM Global
copy /B src\const.js+src\global.js mado-debug.js

REM Util
copy /B mado-debug.js+src\specialfolders.js mado-debug.js

REM Language Extension
copy /B mado-debug.js+src\array.js mado-debug.js
copy /B mado-debug.js+src\date.js mado-debug.js
copy /B mado-debug.js+src\string.js mado-debug.js

REM File
copy /B mado-debug.js+src\file.js mado-debug.js
copy /B mado-debug.js+src\dir.js mado-debug.js
copy /B mado-debug.js+src\drive.js mado-debug.js

REM Network
copy /B mado-debug.js+src\http.js mado-debug.js
copy /B mado-debug.js+src\ftp.js mado-debug.js

REM Other
copy /B mado-debug.js+src\keysender.js mado-debug.js
copy /B mado-debug.js+src\process.js mado-debug.js
copy /B mado-debug.js+src\registry.js mado-debug.js
copy /B mado-debug.js+src\excel.js mado-debug.js
copy /B mado-debug.js+src\clipboardie.js mado-debug.js
copy /B mado-debug.js+src\clipboardexcel.js mado-debug.js
copy /B mado-debug.js+src\clipboard.js mado-debug.js
copy /B mado-debug.js+src\adoconnection.js mado-debug.js
copy /B mado-debug.js+src\adocommand.js mado-debug.js
copy /B mado-debug.js+src\adoaccessconnection.js mado-debug.js
copy /B mado-debug.js+src\adooracleconnection.js mado-debug.js
copy /B mado-debug.js+src\logevent.js mado-debug.js
copy /B mado-debug.js+src\shell.js mado-debug.js

REM jsファイルを圧縮
java -jar yuicompressor-2.4.2\build\yuicompressor-2.4.2.jar --type js -o mado.js mado-debug.js