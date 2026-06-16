' Starts the sync in the background with no visible window.
' Called by Task Scheduler on Windows login.
' Prefers GCalSync.exe; falls back to pythonw + src\sync.py.
Dim fso, shell, pasta, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

pasta = fso.GetParentFolderName(WScript.ScriptFullName)

gCalExe = pasta & "\GCalSync.exe"
If fso.FileExists(gCalExe) Then
    cmd = Chr(34) & gCalExe & Chr(34) & " run"
Else
    pythonw = pasta & "\.venv\Scripts\pythonw.exe"
    If Not fso.FileExists(pythonw) Then
        pythonw = "pythonw.exe"
    End If
    script = pasta & "\src\sync.py"
    cmd = Chr(34) & pythonw & Chr(34) & " " & Chr(34) & script & Chr(34) & " run"
End If

shell.Run cmd, 0, False

Set shell = Nothing
Set fso = Nothing
