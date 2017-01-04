'This script will list all computers on your domain

Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1

Wscript.Echo "The output will be written to C:\FolderLocation\Computers.txt"

strFile = "Computers.txt"
strWritePath = "C:\temp\" & strFile
strDirectory = "C:\temp\"

Set objFSO1 = CreateObject("Scripting.FileSystemObject")

If objFSO1.FileExists(strWritePath) Then
    Set objFolder = objFSO1.GetFile(strWritePath)

Else
    Set objFile = objFSO1.CreateTextFile(strDirectory & strFile)
    objFile = ""

End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set textFile = fso.OpenTextFile(strWritePath, OPEN_FILE_FOR_WRITING)

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Comment from 'LDAP://OU=par,OU=us,DC=sbrinc,DC=com' " _
        & "Where objectClass='computer'"
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    'Wscript.Echo "Computer Name: " & objRecordSet.Fields("Name").Value
    textFile.WriteLine(objRecordSet.Fields("Name").Value)
    objRecordSet.MoveNext
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments
Set objTextFile = objFSO.OpenTextFile(strWritePath, ForReading)

Do Until objTextFile.AtEndOfStream
    strReg = objTextFile.Readline
Loop

WScript.Echo "Complete"
