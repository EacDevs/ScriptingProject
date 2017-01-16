Option Explicit
'Variabelen voor WindowsCheck:
Dim strComputer
Dim objOperatingSystem
Dim objWMIService
Dim colSettings
'Variabelen voor ProgramCheck:
Dim objFSO
Dim objFolder
Dim objFiles
Dim WorkingDir
'Variabelen voor OfficeCheck:
Dim OfficeDir
Dim objFSO2
Dim objFolder2
Dim objFiles2


'Roep de Functies op.
Call WindowsCheck
Call ProgramCheck
Call OfficeCheck

'------------------------------

'Func WindowsCheck
Function WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
MsgBox "OS Name: " & objOperatingSystem.Name, 64
Next
End Function

'------------------------------

'Func ProgramCheck
Function ProgramCheck

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 

WorkingDir = Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& "Software\"
'Als PAS,SAP,Focus niet in dezelfde folder staat als dit script, Waar is PASv3.exe, SAPv2.3.exe en Focusv2.exe dan geïnstalleerd? (uncomment next line)
'WorkingDir = "C:\Program Files(x86)\" 

Set objFolder = objFSO.GetFolder(Left(WorkingDir, Len(WorkingDir)-1))
Set objFiles = objFolder.Files

If objFSO.FileExists(WorkingDir & "PASv3.exe") Then
        MsgBox "PASv3.exe is Geinstalleerd.", 64
	Else
		MsgBox "PASv3.exe mist op het systeem!!!", 16
End If
If objFSO.FileExists(WorkingDir & "SAPv2.3.exe") Then
        MsgBox "SAPv2.3.exe is Geinstalleerd.", 64
	Else
		MsgBox "SAPv2.3.exe mist op het systeem!!!", 16
End If
If objFSO.FileExists(WorkingDir & "Focusv2.exe") Then
        MsgBox "Focusv2.exe is Geinstalleerd.", 64
	Else
		MsgBox "Focusv2.exe mist op het systeem!!!", 16
End If
End Function

Function OfficeCheck
Set objFSO2 = WScript.CreateObject("Scripting.FileSystemObject") 
OfficeDir = "C:\Program Files (x86)\Microsoft Office\root\Office16\" 

Set objFolder2 = objFSO2.GetFolder(Left(OfficeDir, Len(OfficeDir)-1))
Set objFiles2 = objFolder2.Files

If objFSO2.FileExists(OfficeDir & "WINWORD.EXE") Then
        MsgBox "Microsoft Word is Geinstalleerd.", 64
	Else
		MsgBox "Microsoft Word mist op het systeem!!!", 16
End If

If objFSO2.FileExists(OfficeDir & "EXCEL.EXE") Then
        MsgBox "Microsoft Excel is Geinstalleerd.", 64
	Else
		MsgBox "Microsoft Excel mist op het systeem!!!", 16
End If


End Function