var nparams = WScript.Arguments.Named;
var params = WScript.Arguments.Unnamed;

if (!nparams.exists('f')) usage();
if (params.Count != 1) usage();



function usage() {
  var msg = 'main.bat /f [/o] ���t�@�C��\n' +
            '  ���t�@�C��  �ϊ����̃t�@�C��\n' +
            '  /f  �ݒ�t�@�C��\n' +
            '  /o  �o�̓t�@�C����\n';
  WScript.Echo(msg);
  WScript.Quit();
}
