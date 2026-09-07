# sync_startup.ps1 - Synchronise les modeles du poste avec Modeles\Deploy\
# du dossier partage. A planifier a l'OUVERTURE DE SESSION (avant Word).
# -Role : Medecin (Cabinet.dotm -> STARTUP Word), Secretaire (idem +
#         Cabinet.xlsm -> Documents\CabinetCardio), Tous
param(
    [ValidateSet('Medecin', 'Secretaire', 'Tous')][string]$Role = 'Medecin'
)
$ErrorActionPreference = 'Stop'

$cheminTxt = Join-Path $env:APPDATA 'CabinetCardio\chemin.txt'
if (-not (Test-Path $cheminTxt)) {
    Write-Host "ERREUR : $cheminTxt introuvable (lancez d'abord deploy.ps1 sur ce poste)."
    exit 1
}
$racine = (Get-Content $cheminTxt -TotalCount 1).Trim()
$deploy = Join-Path $racine 'Modeles\Deploy'
if (-not (Test-Path $deploy)) { Write-Host "ERREUR : $deploy introuvable."; exit 1 }

function Sync-SiPlusRecent([string]$source, [string]$dest) {
    if (-not (Test-Path $source)) { return }
    $doitCopier = $true
    if (Test-Path $dest) {
        $doitCopier = (Get-Item $source).LastWriteTime -gt (Get-Item $dest).LastWriteTime
    }
    if ($doitCopier) {
        New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
        Copy-Item $source $dest -Force
        Write-Host "mis a jour : $dest"
    } else {
        Write-Host "a jour     : $dest"
    }
}

# Cabinet.dotm dans le demarrage de Word (les 2 roles : la secretaire peut
# aussi ouvrir les courriers dans Word)
$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
Sync-SiPlusRecent (Join-Path $deploy 'Cabinet.dotm') (Join-Path $startup 'Cabinet.dotm')

if ($Role -eq 'Secretaire' -or $Role -eq 'Tous') {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    Sync-SiPlusRecent (Join-Path $deploy 'Cabinet.xlsm') (Join-Path $docs 'CabinetCardio\Cabinet.xlsm')
}
Write-Host 'Synchronisation terminee.'
