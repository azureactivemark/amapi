var xhr = new ActiveXObject ('MSXML2.XMLHTTP');

var url = 'https://www.python.org/ftp/python/2.7.11/python-2.7.11.amd64.msi';
var fileName = 'python-2.7.11.amd64.msi';

xhr.open ('GET', url, false);
xhr.send()

if (xhr.Status == 200) {
  var fso = new ActiveXObject ('Scripting.FileSystemObject');

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
  WScript.Echo ('Installing ' + fileName + ' ...');

  var shell = new ActiveXObject ('WScript.Shell');

  //shell.Run ('msiexec /i ' + fileName + ' /quiet', 1, true);

  WScript.Echo (fileName + 'installed correctly');

  WScript.Echo ('Configuring firewall ...');
  //shell.Run ('netsh advfirewall firewall add rule name="Python27" dir=in action=allow program="C:\\Python27\\python.exe" enable=yes', 1, true);

  WScript.Echo ('Running server ...')
  var file = fso.CreateTextFile ('azure-ws.py');
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

    'server = SimpleXMLRPCServer.SimpleXMLRPCServer (("10.0.0.6", 80))\n' +
    'server.register_function (run)\n' +
    'server.register_function (pyexec)\n' +
    'server.serve_forever()'
  );
  file.Close();

  shell.Run ('C:\\Python27\\python.exe azure-ws.py', 0, false);

  WScript.Echo ("Done!")
}
else {
  WScript.Echo ('Error: HTTP ' + xhr.status + ' ' + xhr.statusText)
}

