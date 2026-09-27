Option Explicit

Dim fso, scriptPath, command, shellApp
Set fso = CreateObject("Scripting.FileSystemObject")
Set shellApp = CreateObject("Shell.Application")

scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), _
    "USB-Drive-Inventory-Collector-GUI.ps1")

If Not fso.FileExists(scriptPath) Then
    MsgBox "The GUI script was not found next to this launcher:" & vbCrLf & scriptPath, _
        vbCritical, "USB Drive Inventory Collector"
    WScript.Quit 1
End If

command = "powershell.exe -NoLogo -NoProfile -STA -File " & _
    Chr(34) & scriptPath & Chr(34)
shellApp.ShellExecute "powershell.exe", Mid(command, Len("powershell.exe") + 2), _
    fso.GetParentFolderName(scriptPath), "runas", 1
