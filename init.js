var xhr = new ActiveXObject ('MSXML2.XMLHTTP');
var shell = new ActiveXObject ('WScript.Shell');
var shApp = new ActiveXObject('Shell.Application')
var fso = new ActiveXObject ('Scripting.FileSystemObject');

// -----------------------------------------------------------------------------
// downloadFile
// -----------------------------------------------------------------------------
function downloadFile (url, fileName) {
  xhr.open ('GET', url, false);
  xhr.send()

  if (xhr.Status == 200) {
    if (fso.FileExists (fileName))
      fso.DeleteFile (fileName);

    WScript.Echo ('Downloading ' + fileName + '...');

    var stream = new ActiveXObject ('ADODB.Stream');
    stream.Open();
    stream.Type = 1 //adTypeBinary
    stream.Write (xhr.ResponseBody);
    stream.Position = 0;
    stream.SaveToFile (fileName);
    stream.Close();

    WScript.Echo (fileName + ' downloaded correctly');

    return true;
  }
  
  WScript.Echo ('Error: HTTP ' + xhr.status + ' ' + xhr.statusText);

  return false;
}

// -----------------------------------------------------------------------------
// installMsi
// -----------------------------------------------------------------------------
function installMsi (fileName) {
  WScript.Echo ('Installing ' + fileName + ' ...');

  shell.Run ('msiexec /i ' + fileName + ' /quiet', 1, true);

  WScript.Echo (fileName + 'installed correctly');
}

// -----------------------------------------------------------------------------
// writeServerScript
// -----------------------------------------------------------------------------
function writeServerScript (fileName) {
  WScript.Echo ('Serializing server script ...');

  var file = fso.CreateTextFile (fileName);
  file.Write (
    'import SimpleXMLRPCServer\n' +
    'import subprocess\n' +

    'def run (cmd) :\n' +
    '  p = subprocess.Popen (\n' +
    '    cmd, \n' +
    '    stdout=subprocess.PIPE,\n' +
    '    stderr=subprocess.PIPE,\n' +
    '    stdin=subprocess.PIPE\n' +
    '  )\n' +
    '  return p.communicate()\n' +

    'def pyexec (cmd) :\n' +
    '  exec (cmd)\n' +
    '  return 0\n' +

    'def test (v) :\n' +
    '  return v\n' +

    'server = SimpleXMLRPCServer.SimpleXMLRPCServer (("10.0.0.6", 80))\n' +
    'server.register_function (run)\n' +
    'server.register_function (pyexec)\n' +
    'server.register_function (test)\n' +
    'server.serve_forever()'
  );
  file.Close();

  WScript.Echo ("Server script serialized");
}

// -----------------------------------------------------------------------------
// unzip
// -----------------------------------------------------------------------------
function unzip (fileName) {
  WScript.Echo ('Uncompressing ' + fileName + ' ...');

  var zip = shApp.NameSpace (fso.getFile (fileName).Path);
  var dst = shApp.NameSpace (fso.getFolder ('.').Path);

  for (var i = 0; i < zip.Items().Count; i++) {
    try {
      WScript.Echo (zip.Items().Item(i));
      dst.CopyHere (zip.Items().Item(i), 4 + 16);
    }
    catch(e) {
      WScript.Echo ('Failed: ' + e);

      return false;
    }
  }

  WScript.Echo (fileName + 'uncompressed correctly');

  return true;
}

// -----------------------------------------------------------------------------
// configureFirewall
// -----------------------------------------------------------------------------
function configureFirewall() {
  WScript.Echo ('Configuring firewall ...');
  shell.Run ('netsh advfirewall firewall add rule name="Python27" dir=in action=allow program="C:\\Python27\\python.exe" enable=yes', 1, true);
  WScript.Echo ('Firewall configured correctly');
}

// -----------------------------------------------------------------------------
// installService
// -----------------------------------------------------------------------------
function installService (nssm, script, serviceName) {
  WScript.Echo ('Installing service ' + serviceName + ' ...');

  var installCmd = nssm + ' install ' + serviceName + ' C:\\Python27\\python.exe ' + script;
  var startCmd = nssm + ' start ' + serviceName;

  shell.Run (installCmd , 1, true);
  shell.Run (startCmd, 1, true);

  WScript.Echo ('Service installed correctly');
}

// -----------------------------------------------------------------------------
// main
// -----------------------------------------------------------------------------
function main() {
  var urlPython = 'https://www.python.org/ftp/python/2.7.11/python-2.7.11.amd64.msi';
  var fileNamePython = 'python-2.7.11.amd64.msi';
  var urlNssm = 'http://nssm.cc/release/nssm-2.24.zip';
  var fileNameNssm = 'nssm-2.24.zip';
  var fileNameScript = 'C:\\azure-ws.py';
  var nssm = "nssm-2.24\\win64\\nssm.exe";

  if (
    downloadFile (urlPython, fileNamePython) &&
    downloadFile (urlNssm, fileNameNssm)
  ) {
    installMsi (fileNamePython);
    unzip (fileNameNssm);
    writeServerScript (fileNameScript);
    configureFirewall();
    installService (nssm, fileNameScript, "AmAzureServicePy");

    WScript.Echo ("Everything ok!");
  }
}


main();
