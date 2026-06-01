Option Explicit

Dim shell, fso, basePath, appPath, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

basePath = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = basePath & "\scripts\product_import_app.py"
shell.CurrentDirectory = basePath

On Error Resume Next
command = "pythonw.exe """ & appPath & """"
shell.Run command, 0, False

If Err.Number <> 0 Then
    Err.Clear
    command = "python.exe """ & appPath & """"
    shell.Run command, 0, False
End If
