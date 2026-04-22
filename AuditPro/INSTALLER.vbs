' ================================================================
' INSTALLER.vbs - Lance l'installation AuditPro
' Double-cliquer sur ce fichier pour installer
' ================================================================

Dim shell, scriptDir, ps1Path, cmd

Set shell   = CreateObject("WScript.Shell")
scriptDir   = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
ps1Path     = scriptDir & "installer_auditpro.ps1"

' Lancer PowerShell en contournant les restrictions
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1Path & """"

' Lancer avec l'interface utilisateur (pas en silencieux)
shell.Run cmd, 1, True

Set shell = Nothing
WScript.Quit 0
