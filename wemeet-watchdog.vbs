' wemeet-watchdog.vbs
' Silent launcher (no console window) for wemeet-watchdog.ps1.
' Resolves the ps1 path from its own location to avoid non-ASCII literals.
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("Wscript.Shell")
ps1 = fso.GetParentFolderName(WScript.ScriptFullName) & "\wemeet-watchdog.ps1"
sh.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & ps1 & """", 0, False
