# ================================================================
# AuditPro - Script d'installation et deploiement
# Lance depuis PowerShell : Set-ExecutionPolicy Bypass; .\installer_auditpro.ps1
# ================================================================

$ErrorActionPreference = "Continue"
$AppName    = "AuditPro"
$AppVersion = "1.0.0"
$Publisher  = "FIDAROC GRANT THORNTON"

# Chemins
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$DistDir    = Join-Path $ScriptDir "dist\AuditPro"
$ExePath    = Join-Path $DistDir "AuditPro.exe"
$TargetDir  = "C:\AuditPro"
$TargetExe  = Join-Path $TargetDir "AuditPro.exe"
$Desktop    = [Environment]::GetFolderPath("Desktop")
$StartMenu  = [Environment]::GetFolderPath("StartMenu")

Write-Host ""
Write-Host "  =================================================="
Write-Host "   AuditPro - Installation"
Write-Host "  =================================================="
Write-Host ""

# ── Étape 1 : Debloquer TOUS les fichiers ───────────────────────
Write-Host "[1/4] Deblocage des fichiers Windows..."
Get-ChildItem -Path $ScriptDir -Recurse -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

if (Test-Path $DistDir) {
    Get-ChildItem -Path $DistDir -Recurse -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
}
Write-Host "       OK"

# ── Étape 2 : Verifier que l'exe existe ─────────────────────────
Write-Host "[2/4] Verification de l'executable..."
if (-not (Test-Path $ExePath)) {
    Write-Host ""
    Write-Host "  [!] AuditPro.exe non trouve dans dist\AuditPro\"
    Write-Host "      Lancement du build PyInstaller..."
    Write-Host ""

    Set-Location $ScriptDir
    & python -m PyInstaller AuditPro.spec --noconfirm --clean

    if (-not (Test-Path $ExePath)) {
        Write-Host "  [ERREUR] Build echoue. Verifiez Python et PyInstaller."
        Read-Host "Appuyez sur Entree pour quitter"
        exit 1
    }

    # Copier les données
    $DataSrc = Join-Path $ScriptDir "data"
    $DataDst = Join-Path $DistDir "data"
    if (Test-Path $DataSrc) {
        Copy-Item -Path $DataSrc -Destination $DataDst -Recurse -Force
    }
}
Write-Host "       OK - $ExePath"

# ── Étape 3 : Copier vers C:\AuditPro (hors Downloads) ──────────
Write-Host "[3/4] Installation dans C:\AuditPro..."

if (Test-Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
Copy-Item -Path "$DistDir\*" -Destination $TargetDir -Recurse -Force

# Debloquer les fichiers copiés
Get-ChildItem -Path $TargetDir -Recurse | Unblock-File -ErrorAction SilentlyContinue

Write-Host "       OK - Installe dans $TargetDir"

# ── Étape 4 : Créer les raccourcis ──────────────────────────────
Write-Host "[4/4] Creation des raccourcis..."

$WshShell = New-Object -ComObject WScript.Shell

# Raccourci Bureau
$ShortcutDesktop = $WshShell.CreateShortcut("$Desktop\AuditPro.lnk")
$ShortcutDesktop.TargetPath       = $TargetExe
$ShortcutDesktop.WorkingDirectory = $TargetDir
$ShortcutDesktop.Description      = "AuditPro - Assistant d'audit FIDAROC GT"
if (Test-Path "$TargetDir\_internal") {
    $ShortcutDesktop.IconLocation = "$TargetExe,0"
}
$ShortcutDesktop.Save()
Write-Host "       Raccourci Bureau cree"

# Raccourci Menu Démarrer
$StartMenuFolder = Join-Path $StartMenu "Programs\AuditPro"
New-Item -ItemType Directory -Path $StartMenuFolder -Force | Out-Null
$ShortcutStart = $WshShell.CreateShortcut("$StartMenuFolder\AuditPro.lnk")
$ShortcutStart.TargetPath       = $TargetExe
$ShortcutStart.WorkingDirectory = $TargetDir
$ShortcutStart.Description      = "AuditPro - Assistant d'audit FIDAROC GT"
$ShortcutStart.IconLocation     = "$TargetExe,0"
$ShortcutStart.Save()
Write-Host "       Raccourci Menu Demarrer cree"

# ── Résumé ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "  =================================================="
Write-Host "   INSTALLATION TERMINEE !"
Write-Host "  =================================================="
Write-Host ""
Write-Host "   Application installee dans : $TargetDir"
Write-Host "   Raccourci Bureau           : OK"
Write-Host "   Raccourci Menu Demarrer    : OK"
Write-Host ""
Write-Host "   POUR LANCER L'APPLICATION :"
Write-Host "   -> Double-cliquez sur AuditPro (Bureau)"
Write-Host "   -> Ou cherchez AuditPro dans le menu Demarrer"
Write-Host ""

# Proposer de lancer immédiatement
$launch = Read-Host "Lancer AuditPro maintenant ? (O/n)"
if ($launch -ne "n" -and $launch -ne "N") {
    Write-Host "  Lancement..."
    Start-Process $TargetExe
}

Write-Host ""
Read-Host "Appuyez sur Entree pour fermer"
