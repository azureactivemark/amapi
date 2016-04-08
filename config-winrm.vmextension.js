var objService = GetObject ('winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\default');
var objReg = objService.Get ('StdRegProv');
var objShell = new ActiveXObject ('WScript.Shell');

var objMethod = objReg.Methods_.Item('EnumKey'); 
var objParamsIn = objMethod.InParameters.SpawnInstance_(); 
objParamsIn.hDefKey = 0x80000002; 
objParamsIn.sSubKeyName = 'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles'; 

var objParamsOut = objReg.ExecMethod_ (objMethod.Name, objParamsIn); 

if (objParamsOut.ReturnValue === 0) {
  if (objParamsOut.sNames != null) {
    var a = objParamsOut.sNames.toArray();
    for (var i = 0; i < a.length; ++i) {
      var key = 'HKLM\\' + objParamsIn.sSubKeyName + '\\' + a[i] + '\\Category';
      objShell.RegWrite (key,  2, 'REG_DWORD'); 
    }
  }
}

objShell.Run ('winrm quickconfig -quiet', 1, true);
objShell.Run ('winrm set winrm/config/service/auth @{Basic="true"}', 1, true);
objShell.Run ('winrm set winrm/config/service @{AllowUnencrypted="true"}', 1, true);
objShell.Run ('netsh advfirewall firewall add rule name="WinRM-HTTP" dir=in localport=5985 protocol=TCP action=allow', 1, true);

shell.Popup ('Done!');