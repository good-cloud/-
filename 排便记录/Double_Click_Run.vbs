Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

folder = fso.GetParentFolderName(WScript.ScriptFullName)
script = fso.BuildPath(folder, "bowel_record_tool.py")
bundledPythonw = "C:\Users\Eva\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"

If fso.FileExists(bundledPythonw) Then
    pythonw = bundledPythonw
Else
    pythonw = "pythonw.exe"
End If

shell.CurrentDirectory = folder
shell.Run Chr(34) & pythonw & Chr(34) & " " & Chr(34) & script & Chr(34), 1, False
