Option Explicit
'VariabelenA voor het txt-bestand
Dim objFSO, wshNetwork, Log, objShell, savedPath
Dim objOperatingSystem, objWMIService, colSettings
'VariabelenB voor CheckRegistry
Dim strComputer, strKey, strEntry1a, strEntry1b, strEntry3, strEntry4, objReg, arrSubkeys, strSubkey, strValue1, intRet1, intValue3, _
intValue4, IgnoreListEntries, i
'VariabelenC voor IgnoreList + Aantal Ignores in totaal.
Dim IgnoreListCounter
Dim strIgnoreList(9)
IgnoreListEntries = 9

'Selecteer de Juiste Keys in REGISTRY
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE 
strComputer = "." 
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
strEntry1a = "DisplayName" 
strEntry1b = "QuietDisplayName" 
strEntry3 = "VersionMajor" 
strEntry4 = "VersionMinor" 

'Roep de verschillende functies op.
Call WriteText
Call WindowsCheck
Call BuildIgnore
Call CheckRegistry
Call CloseText

'Functie voor het Maken van het textbestand.
Function WriteText
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
Set wshNetwork = CreateObject( "WScript.Network" )

Set Log = objFSO.CreateTextFile(Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt", 2)
End Function

'Functie voor het checken van de windowsversie.
Function WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
MsgBox "OS Name: " & objOperatingSystem.Name, 64 
Next
End Function

'Specifieer hier de programma's die genegeerd moeten worden. Vergeet niet om de twee getallen in VariabelenC aan te passen.
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


'Hier wordt de Registry geraadpleegd.
Function CheckRegistry
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv") 
objReg.EnumKey HKLM, strKey, arrSubkeys 
Log.WriteLine "Installed Illegal Applications:" & VbCrLf 

'--------------------
'Haal voor iedere key die je vindt de Display Name en Versie Nummer op
For Each strSubkey In arrSubkeys 
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1) 
  If intRet1 <> 0 Then 
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1 
  End If
  
	objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
	objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
'--------------------


  	If strValue1 <> "" Then
		IgnoreListCounter = 0
		Do While IgnoreListCounter < IgnoreListEntries
			If strValue1 = strIgnoreList(IgnoreListCounter) Then
			
				If intValue3 <> "" then
					MsgBox strValue1 & " is geinstalleerd met versie nummer " & intValue3 & "." & intValue4, 64
				Else
					MsgBox strValue1 & " is geinstalleerd.", 64
				End If
			Exit do
			End If
			
			IgnoreListCounter = IgnoreListCounter + 1
		Loop
			
			If IgnoreListCounter = IgnoreListEntries Then			
				If strValue1 <> strIgnoreList(IgnoreListCounter) Then
					Log.WriteLine VbCrLf & "Display Name: " & strValue1
					If intValue3 <> "" Then 
					Log.WriteLine "Version: " & intValue3 & "." & intValue4 
					End If
				End If
				If strValue1 = strIgnoreList(IgnoreListCounter) Then
			End If
		End If
	End If
  Next

End Function

'Sluit het textbestand en vraag de gebruiker of hij deze wil openen.
Function CloseText
Set objShell = CreateObject("WScript.Shell")
savedPath = Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt"
MsgBox "De Informatie over de Illegale Programma's zijn weggeschreven naar " & savedPath, 64
Log.Close
If MsgBox("Wilt u het bestand met de informatie nu openen?", vbYesNo) = vbYes then
	objShell.Run("notepad.exe " & Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullname) - Len(Wscript.ScriptName))& wshNetwork.ComputerName & ".txt")
End If
End Function