Option Explicit
'Variabelen voor WindowsCheck:
Dim strComputer
Dim objOperatingSystem
Dim objWMIService
Dim colSettings

Call setVariabelen
Call WindowsCheck



Function setVariabelen
'WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
End Function


'WindowsCheck
Function WindowsCheck
For Each objOperatingSystem in colSettings 
Wscript.Echo "OS Name: " & objOperatingSystem.Name
Next
End Function
