Option Explicit
Dim ows: Set ows = WScript.CreateObject("WScript.Shell")
Dim osc: Set osc = ows.CreateShortcut(ows.SpecialFolders("SENDTO") & "\MSIDump.lnk")
osc.TargetPath = "C:\BIN\msidump.exe"
osc.Save 

Set osc = Nothing
Set ows = Nothing