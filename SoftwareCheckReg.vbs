Option Explicit
Dim strComputer, strKey, strEntry1a, strEntry1b, strEntry3, strEntry4, objReg, arrSubkeys, strSubkey, strValue1, intRet1, intValue3, intValue4
Dim objFSO, wshNetwork, Log

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
strComputer = "." 
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
strEntry1a = "DisplayName" 
strEntry1b = "QuietDisplayName" 
strEntry3 = "VersionMajor" 
strEntry4 = "VersionMinor" 

Call WriteText
Call CheckRegistry
Call CloseText

'
Function WriteText
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
Set wshNetwork = CreateObject( "WScript.Network" )

Set Log = objFSO.CreateTextFile(Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 2)
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
  If strValue1 <> "" Then 
    Log.WriteLine VbCrLf & "Display Name: " & strValue1 
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
  objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
  If intValue3 <> "" Then 
     Log.WriteLine "Version: " & intValue3 & "." & intValue4 
  End If 
Next
End Function

Function CloseText
MsgBox "De Informatie over het object " & wshNetwork.ComputerName & " is weggeschreven naar " & Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 64
Log.Close
End Function