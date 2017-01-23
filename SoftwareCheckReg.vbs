Option Explicit
'Variabelen voor het txt-bestand
Dim objFSO, wshNetwork, Log
'Variabelen voor CheckRegistry
Dim strComputer, strKey, strEntry1a, strEntry1b, strEntry3, strEntry4, objReg, arrSubkeys, strSubkey, strValue1, intRet1, intValue3, _
intValue4, strIgnoreList(8), IgnoreListCounter, i
Const IgnoreListEntries = 12

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
strComputer = "." 
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
strEntry1a = "DisplayName" 
strEntry1b = "QuietDisplayName" 
strEntry3 = "VersionMajor" 
strEntry4 = "VersionMinor" 

Call WriteText
Call BuildIgnore
Call CheckRegistry
Call CloseText

Function WriteText
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
Set wshNetwork = CreateObject( "WScript.Network" )

Set Log = objFSO.CreateTextFile(Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 2)
End Function

Function BuildIgnore
strIgnoreList(0) = "Windows 10"
strIgnoreList(1) = "Microsoft Word 2016"
strIgnoreList(2) = "Microsoft Excel 2016"
strIgnoreList(3) = "Microsoft Office"
strIgnoreList(4) = "Microsoft Office 365 ProPlus - en-us"
strIgnoreList(5) = "Microsoft Office 365 ProPlus - nl-nl"
strIgnoreList(6) = "P.A.S."
strIgnoreList(7) = "Focus-agenda"
strIgnoreList(8) = "SAP"
End Function

Function CheckRegistry
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 
objReg.EnumKey HKLM, strKey, arrSubkeys 
Log.WriteLine "Installed Applications:" & VbCrLf 
For Each strSubkey In arrSubkeys 
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1) 
  If intRet1 <> 0 Then 
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1 
  End If
  For i = 0 to 8
  	If strValue1 <> "" Then
	IgnoreListCounter = 0
	End If
	If strValue1 <> strIgnoreList(i) Then 
		Log.WriteLine VbCrLf & "Display Name: " & strValue1 
		
		objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
		objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
		If intValue3 <> "" Then 
			Log.WriteLine "Version: " & intValue3 & "." & intValue4 
		End If
	End If
  Next
Next
End Function

Function CloseText
MsgBox "De Informatie over het object " & wshNetwork.ComputerName & " is weggeschreven naar " & Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 64
Log.Close
End Function