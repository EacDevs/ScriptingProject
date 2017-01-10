Option Explicit
Dim strComputer
Dim objOperatingSystem
Dim objWMIService
Dim colSettings

Call WindowsCheck

Function WindowsCheck
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSettings = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colSettings 
Wscript.Echo "OS Name: " & objOperatingSystem.Name
Next
End Function