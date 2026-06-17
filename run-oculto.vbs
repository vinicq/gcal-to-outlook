' Starts the app on Windows login without a console window.
' Opens the tray monitor in "tray" mode: it shows the tray icon AND drives the
' sync loop itself, so a single launch covers both visibility and syncing.
' Prefers GCalSync.exe; falls back to pythonw + src\app.py.
Dim fso, shell, pasta, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

pasta = fso.GetParentFolderName(WScript.ScriptFullName)

gCalExe = pasta & "\GCalSync.exe"
If fso.FileExists(gCalExe) Then
    cmd = Chr(34) & gCalExe & Chr(34) & " tray"
Else
    pythonw = pasta & "\.venv\Scripts\pythonw.exe"
    If Not fso.FileExists(pythonw) Then
        pythonw = "pythonw.exe"
    End If
    script = pasta & "\src\app.py"
    cmd = Chr(34) & pythonw & Chr(34) & " " & Chr(34) & script & Chr(34) & " tray"
End If

shell.Run cmd, 0, False

Set shell = Nothing
Set fso = Nothing
