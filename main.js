var nparams = WScript.Arguments.Named;
var params = WScript.Arguments.Unnamed;

if (nparams.Exists('h')) usage();
if (params.Count != 1) usage();

var fileSysObj = new ActiveXObject('Scripting.FileSystemObject');
var scriptPath = String(WScript.ScriptFullName).replace(WScript.ScriptName, '');

var configName = 'config.json';
if (nparams.Exists('f')) {
  configName = nparams.Item('f');
}

var outputFileName = 'output.pptx';
if (nparams.Exists('o')) {
  outputFileName = nparams.Item('o');
}

var configPath = fileSysObj.BuildPath(scriptPath, configName);
var srcPath    = fileSysObj.BuildPath(scriptPath, params.Item(0));
var dstPath    = fileSysObj.BuildPath(scriptPath, outputFileName);

fileSysObj.CopyFile(srcPath, dstPath);


// load config file
var configFile = fileSysObj.OpenTextFile(configPath, 1, false, -2);
var config = eval('(' + configFile.ReadAll() + ')');
// descending sort
config.del.sort(function(a, b) {
  return b - a;
});


// edit powerpoint file
var pptObj = WScript.CreateObject('PowerPoint.Application');
pptObj.Presentations.Open(dstPath);

for (var i = 0; i < config.del.length; i++) {
  pptObj.ActivePresentation.Slides(config.del[i]).Delete();
}

pptObj.ActivePresentation.Save();
pptObj.Quit();


function usage() {
  var msg = 'main.bat [/f] [/o] ���t�@�C��\n' +
            '  ���t�@�C��  �ϊ����̃t�@�C��\n' +
            '  /f  �ݒ�t�@�C�� (default:config.json)\n' +
            '  /o  �o�̓t�@�C���� (default:output.pptx)\n' +
            '  /h  �w���v\n';
  WScript.Echo(msg);
  WScript.Quit();
}
