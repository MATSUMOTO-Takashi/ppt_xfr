var nparams = WScript.Arguments.Named;
var params = WScript.Arguments.Unnamed;

if (!nparams.exists('f')) usage();
if (params.Count != 1) usage();



function usage() {
  var msg = 'main.bat /f [/o] 元ファイル\n' +
            '  元ファイル  変換元のファイル\n' +
            '  /f  設定ファイル\n' +
            '  /o  出力ファイル名\n';
  WScript.Echo(msg);
  WScript.Quit();
}
