Option Explicit

' Launches markitdown-convert.ps1 with no visible console windows.
' Invoked from Explorer: wscript.exe "...\markitdown-convert-launcher.vbs" "%1"
' Window hiding is done by WshShell.Run (style 0), not by starting powershell visibly.

If WScript.Arguments.Count < 1 Then
  WScript.Quit 1
End If

Dim fso, scriptDir, ps1Path, shell, cmdLine, exitCode
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1Path = scriptDir & "\markitdown-convert.ps1"

If Not fso.FileExists(ps1Path) Then
  WScript.Quit 1
End If

' 0 = hidden window for the child process; True = wait for completion.
cmdLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
          Chr(34) & ps1Path & Chr(34) & " " & Chr(34) & WScript.Arguments(0) & Chr(34)

exitCode = shell.Run(cmdLine, 0, True)
WScript.Quit exitCode
