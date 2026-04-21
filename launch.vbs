' Silent launcher — runs convert.ps1 with zero window visibility
' This avoids the brief PowerShell console flash that -WindowStyle Hidden still shows
Set shell = CreateObject("WScript.Shell")
arg = Chr(34) & WScript.Arguments(0) & Chr(34)
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
ps1 = Chr(34) & scriptDir & "\convert.ps1" & Chr(34)
cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File " & ps1 & " " & arg
shell.Run cmd, 0, False
