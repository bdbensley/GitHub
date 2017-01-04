Set objWshShell = WScript.CreateObject("WScript.Shell")
objWshShell.Popup "Connecting to Network." & vbCrLf & "Please wait...", 3, "Wait...",0
Set objWshShell = nothing
do Until Pingable("PARFSCL02") 'Can be host name or ip address
Wscript.Sleep(5000) 'Pause 5 Seconds
Loop
wscript.echo "Connected to Network"
msgbox "Insert your code here!"
Function Pingable(strHost) 
    Const OpenAsASCII = 0 
    Const FailIfNotExist = 0 
    Const ForReading =  1 
    Dim objShell
	Dim objFSO
	Dim strTempFle
	Dim objFile 
    Set objShell = CreateObject("WScript.Shell") 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    strTempFle = objFSO.GetSpecialFolder(2).ShortPath & "\" & objFSO.GetTempName 
    objShell.Run "%comspec% /c ping.exe -n 2 -w 500 " & strHost & ">" & strTempFle, 0 , True 
    Set objFile = objFSO.OpenTextFile(strTempFle, ForReading, FailIfNotExist, OpenAsASCII) 
    Select Case InStr(objFile.ReadAll, "TTL=") 
        Case 0
            Pingable = False 
        Case Else
            Pingable = True 
    End Select 
    objFile.Close 
    objFSO.DeleteFile(strTempFle)
End Function 

