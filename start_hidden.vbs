' ── Held Time Report — hidden launcher ──────────────────────────────────────
' Runs start.bat silently in the background (no console window).
' This is what Task Scheduler calls at login.

Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run Chr(34) & WScript.ScriptFullName & Chr(34), 0, False

Dim bat
bat = Replace(WScript.ScriptFullName, "start_hidden.vbs", "start.bat")
shell.Run "cmd /c " & Chr(34) & bat & Chr(34), 0, False

Set shell = Nothing
