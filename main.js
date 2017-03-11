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

var tagName = null;
if (nparams.Exists('t')) {
  tagName = nparams.Item('t');
}

var outputFileName = 'output.pptx';
if (nparams.Exists('o')) {
  outputFileName = nparams.Item('o');
}

var configPath = fileSysObj.BuildPath(scriptPath, configName);
var srcPath    = fileSysObj.BuildPath(scriptPath, params.Item(0));
var dstPath    = fileSysObj.BuildPath(scriptPath, outputFileName);

fileSysObj.CopyFile(srcPath, dstPath);

// edit powerpoint file
var pptObj = WScript.CreateObject('PowerPoint.Application');
pptObj.Presentations.Open(dstPath);


// drop slides with tag name
if (tagName !== null) {
  dropWithTag(tagName, pptObj);

  pptObj.ActivePresentation.Save();
  pptObj.Quit();

  WScript.Quit();
}


// load config file
var configFile = fileSysObj.OpenTextFile(configPath, 1, false, -2);
var config = eval('(' + configFile.ReadAll() + ')');
// descending sort
config.slide.del.sort(function(a, b) {
  return b - a;
});

// note
if (config.note.del === -1) {
  // slide number is 1 origin
  for (var i = 1; i <= pptObj.ActivePresentation.Slides.Count; i++) {
    pptObj.ActivePresentation.Slides(i).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange = '';
  }
} else {
  for (var i = 0; i < config.note.del.length; i++) {
    pptObj.ActivePresentation.Slides(config.note.del[i]).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange = '';
  }
}

// slide
if (config.slide.show === -1) {
  // slide number is 1 origin
  for (var i = 1; i <= pptObj.ActivePresentation.Slides.Count; i++) {
    pptObj.ActivePresentation.Slides(i).SlideShowTransition.Hidden = false;
  }
} else {
  for (var i = 0; i < config.slide.show.length; i++) {
    pptObj.ActivePresentation.Slides(config.slide.show[i]).SlideShowTransition.Hidden = false;
  }
}

for (var i = 0; i < config.slide.hidden.length; i++) {
  pptObj.ActivePresentation.Slides(config.slide.hidden[i]).SlideShowTransition.Hidden = true;
}

for (var i = 0; i < config.slide.del.length; i++) {
  pptObj.ActivePresentation.Slides(config.slide.del[i]).Delete();
}

pptObj.ActivePresentation.Save();
pptObj.Quit();


function dropWithTag(tag, pptObj) {
  tag += '\r';
  for (var i = 1; i <= pptObj.ActivePresentation.Slides.Count; i++) {
    if (String(pptObj.ActivePresentation.Slides(i).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange).slice(0, tag.length) === tag) {
      pptObj.ActivePresentation.Slides(i).Delete();
      i--;
    }
  }
}

function usage() {
  var msg = 'main.bat [/f] [/o] 元ファイル\n' +
            'main.bat [/t] [/o] 元ファイル\n' +
            '  元ファイル  変換元のファイル\n' +
            '  /f  設定ファイル (default:config.json)\n' +
            '  /t  削除対象のタグ名\n' +
            '  /o  出力ファイル名 (default:output.pptx)\n' +
            '  /h  ヘルプ\n';
  WScript.Echo(msg);
  WScript.Quit();
}
