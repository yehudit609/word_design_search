' WordFormatSearch - starts node server.js with no visible window.
' Usage: wscript.exe start-server.vbs "C:\Program Files\nodejs\node.exe"
Option Explicit
Dim sh, fso, dir, node, rc
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
If WScript.Arguments.Count > 0 Then
  node = WScript.Arguments(0)
Else
  node = "node.exe"
End If
sh.CurrentDirectory = dir
' 0 = hidden window, True = wait (so the scheduled task stays "Running")
rc = sh.Run("""" & node & """ """ & dir & "\server.js""", 0, True)
WScript.Quit rc
