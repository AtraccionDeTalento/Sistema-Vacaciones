Dim shell, fso, rootDir, runnerPath, installerPath

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

rootDir = fso.GetParentFolderName(WScript.ScriptFullName)
runnerPath = rootDir & "\run_sistema.bat"
installerPath = rootDir & "\INSTALAR_SISTEMA_VACACIONES.bat"

If Not fso.FileExists(rootDir & "\.venv\Scripts\python.exe") Then
    MsgBox "No se encontro el entorno listo. Primero se ejecutara el instalador.", 64, "Sistema de Vacaciones USIL"
    If fso.FileExists(installerPath) Then
        shell.Run "cmd.exe /c """ & installerPath & """", 1, True
    Else
        MsgBox "No se encontro el instalador en: " & rootDir, 16, "Sistema de Vacaciones USIL"
        WScript.Quit 1
    End If
End If

If Not fso.FileExists(runnerPath) Then
    MsgBox "No se encontro run_sistema.bat en: " & rootDir, 16, "Sistema de Vacaciones USIL"
    WScript.Quit 1
End If

shell.Run "cmd.exe /c """ & runnerPath & """", 0, False