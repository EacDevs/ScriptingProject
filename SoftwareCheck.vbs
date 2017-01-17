Option Explicit
'Algemene Variabelen
Dim wshNetwork
Dim Log
Dim objFSO
'Variabelen voor WindowsCheck:
Dim strComputer
Dim objOperatingSystem
Dim objWMIService
Dim colSettings
'Variabelen voor ProgramCheck:
Dim WorkingDir
Dim objFolder
Dim objFiles
'Variabelen voor OfficeCheck:
Dim OfficeDir
Dim objFolder2
Dim objFiles2


'Roep de Functies op.
Call WriteText
Call WindowsCheck
Call ProgramCheck
Call OfficeCheck
Call Illegal
Call CloseText

'------------------------------

'Func WriteText
Function WriteText
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
Set wshNetwork = CreateObject( "WScript.Network" )

Set Log = objFSO.CreateTextFile(Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 2)
End Function

'Func WindowsCheck
Function WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
Log.WriteLine "OS Name: " & objOperatingSystem.Name
Next
Log.WriteLine "----------------------------------------------------------------"
End Function

'------------------------------

'Func ProgramCheck
Function ProgramCheck
WorkingDir = Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& "Software\"
'Als PAS,SAP,Focus niet in dezelfde folder staat als dit script, Waar is PASv3.exe, SAPv2.3.exe en Focusv2.exe dan geïnstalleerd? (uncomment next line)
'WorkingDir = "C:\Program Files(x86)\" 

Set objFolder = objFSO.GetFolder(Left(WorkingDir, Len(WorkingDir)-1))
Set objFiles = objFolder.Files

If objFSO.FileExists(WorkingDir & "PASv3.exe") Then
        Log.WriteLine "PASv3.exe is Geinstalleerd."
	Else
		Log.WriteLine "PASv3.exe mist op het systeem!!!"
End If
If objFSO.FileExists(WorkingDir & "SAPv2.3.exe") Then
        Log.WriteLine "SAPv2.3.exe is Geinstalleerd."
	Else
		Log.WriteLine "SAPv2.3.exe mist op het systeem!!!"
End If
If objFSO.FileExists(WorkingDir & "Focusv2.exe") Then
        Log.WriteLine "Focusv2.exe is Geinstalleerd."
	Else
		Log.WriteLine "Focusv2.exe mist op het systeem!!!"
End If
End Function

Function OfficeCheck
OfficeDir = "C:\Program Files (x86)\Microsoft Office\root\Office16\" 

Set objFolder2 = objFSO.GetFolder(Left(OfficeDir, Len(OfficeDir)-1))
Set objFiles2 = objFolder2.Files

If objFSO.FileExists(OfficeDir & "WINWORD.EXE") Then
        Log.WriteLine "Microsoft Word is Geinstalleerd."
	Else
		Log.WriteLine "Microsoft Word mist op het systeem!!!"
End If

If objFSO.FileExists(OfficeDir & "EXCEL.EXE") Then
        Log.WriteLine "Microsoft Excel is Geinstalleerd."
	Else
		Log.WriteLine "Microsoft Excel mist op het systeem!!!"
End If

Log.WriteLine "----------------------------------------------------------------"
Log.WriteLine "De Volgende illegale Office Programma's zijn gevonden:"
Log.WriteLine "----------------------------------------------------------------"
If objFSO.FileExists(OfficeDir & "GROOVE.EXE") Then
		Log.WriteLine "Onedrive is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "lync.exe") Then
		Log.WriteLine "Skype for Business is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "MSACCESS.EXE") Then
		Log.WriteLine "Microsof Access is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "MSPUB.EXE") Then
		Log.WriteLine "Microsoft Publisher is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "ONENOTE.EXE") Then
		Log.WriteLine "Microsoft OneNote is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "OUTLOOK.EXE") Then
		Log.WriteLine "Microsoft Outlook is Geinstalleerd!!!"
End If

If objFSO.FileExists(OfficeDir & "VISIO.EXE") Then
		Log.WriteLine "Microsoft Visio is Geinstalleerd!!!"
End If
End Function

Function Illegal
Log.WriteLine "----------------------------------------------------------------"
Log.WriteLine "De Volgende Programma's zijn gevonden in C:\Program Files (x86):"
Log.WriteLine "----------------------------------------------------------------"
For Each objFolder in objFSO.GetFolder("C:\Program Files (x86)\").SubFolders
Log.WriteLine objFolder.Path
Next
Log.WriteLine ""
Log.WriteLine "----------------------------------------------------------------"
Log.WriteLine "De Volgende Programma's zijn gevonden in C:\Program Files:"
Log.WriteLine "----------------------------------------------------------------"
For Each objFolder in objFSO.GetFolder("C:\Program Files\").SubFolders
Log.WriteLine objFolder.Path
Next

End Function

Function CloseText
MsgBox "De Informatie over het object " & wshNetwork.ComputerName & " is weggeschreven naar " & Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 64
Log.Close
End Function