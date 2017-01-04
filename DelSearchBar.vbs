Option Explicit
Dim SysVarReg, Value
Set SysVarReg = WScript.CreateObject("WScript.Shell")
SysVarReg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarMode", 0