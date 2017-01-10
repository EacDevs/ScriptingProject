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


'Roep de Functies op.
Call WindowsCheck
Call ProgramCheck

'------------------------------

'Func WindowsCheck
Function WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
Wscript.Echo "OS Name: " & objOperatingSystem.Name
Next
End Function

'------------------------------

'Func ProgramCheck
Function ProgramCheck

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 

'Waar is PASv3.exe etc geïnstalleerd?
WorkingDir = "D:\School\Blok6\Scripting\ScriptingProject\" 
Set objFolder = objFSO.GetFolder((Left(WorkingDir, Len(WorkingDir)-1)))
Set objFiles = objFolder.Files

If objFSO.FileExists(WorkingDir & "PASv3.exe") Then
        MsgBox "PASv3.exe is Geinstalleerd."
	Else
		MsgBox "PASv3.exe mist op het systeem!!!", 48
End If
If objFSO.FileExists(WorkingDir & "SAPv2.3.exe") Then
        MsgBox "SAPv2.3.exe is Geinstalleerd."
	Else
		MsgBox "SAPv2.3.exe mist op het systeem!!!", 48
End If
End Function
