# installer_cabinet.ps1 - Installation COMPLETE du logiciel de cabinet sur un
# poste, a partir du paquet d'installation (dossier "install" du NAS).
#
# Le paquet contient :
#   Build\installer_cabinet.ps1 (ce script), sync_startup.ps1, installer_relais.ps1, init_donnees.ps1
#   Src\Word\Relais\modCabinetRelais.bas   (raccourcis clavier, poste medecin)
#   Donnees\   (bases patients + correspondants, config, nomenclature, modeles construits)
#
# POSTE SECRETAIRE (a faire en premier) - la racine des donnees est creee sur ce
# poste puis partagee sur le reseau :
#   .\installer_cabinet.ps1 -Role Secretaire -Racine "C:\CabinetCardio"
# POSTE MEDECIN - la racine est le partage du poste secretaire :
#   .\installer_cabinet.ps1 -Role Medecin -Racine "\\POSTE-SECRETAIRE\CabinetCardio"
#
# Relancable sans risque : ne remplace jamais une base deja presente.
param(
    [Parameter(Mandatory = $true)][ValidateSet('Medecin', 'Secretaire')][string]$Role,
    [Parameter(Mandatory = $true)][string]$Racine,
    [string]$NomPartage = 'CabinetCardio',
    [switch]$SansRelais
)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

$paquet = Split-Path $PSScriptRoot -Parent
$seed = Join-Path $paquet 'Donnees'
$erreurs = @()
function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t) { Write-Host "  OK  $t" -ForegroundColor Green }
function Ko([string]$t) { Write-Host "  !!  $t" -ForegroundColor Red; $script:erreurs += $t }

# ---------------------------------------------------------------- 0. controles
Etape 'Controles prealables'
if (-not (Test-Path $seed)) { throw "Paquet incomplet : dossier Donnees absent ($seed)" }
foreach ($p in 'Word.Application', 'Excel.Application') {
    try { $o = New-Object -ComObject $p; $o.Quit(); Ok "$p disponible" } catch { Ko "$p indisponible : Office 2016 installe ?" }
}
$wordOuvert = Get-Process WINWORD -ErrorAction SilentlyContinue
$excelOuvert = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($wordOuvert -or $excelOuvert) {
    Write-Host ''
    Write-Host 'Word et/ou Excel sont ouverts. Fermez-les COMPLETEMENT (y compris les fenetres reduites) puis relancez.' -ForegroundColor Yellow
    exit 1
}
if ($Role -eq 'Medecin' -and -not (Test-Path $Racine)) {
    throw "Racine inaccessible : $Racine - le poste secretaire est-il allume et le dossier partage ?"
}

# ---------------------------------------------------------------- 1. donnees
Etape "Donnees du cabinet : $Racine"
if ($Role -eq 'Secretaire') {
    New-Item -ItemType Directory -Force -Path $Racine | Out-Null
    # copie de tout ce qui n'existe pas encore (jamais d'ecrasement d'une base)
    $n = 0
    Get-ChildItem $seed -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($seed.Length + 1)
        $dest = Join-Path $Racine $rel
        New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
        if (-not (Test-Path $dest)) { Copy-Item $_.FullName $dest; $n++ }
    }
    foreach ($d in 'Base\locks', 'Actes', 'Patients', 'Echange\AEnvoyer', 'Echange\Traites', 'Sauvegardes', 'Logs') {
        New-Item -ItemType Directory -Force -Path (Join-Path $Racine $d) | Out-Null
    }
    Ok "$n fichier(s) copie(s) (les fichiers deja presents sont conserves)"
    # les modeles construits sont toujours mis a jour (version du paquet)
    Copy-Item (Join-Path $seed 'Modeles\Deploy\*') (Join-Path $Racine 'Modeles\Deploy\') -Force
    Copy-Item (Join-Path $PSScriptRoot 'sync_startup.ps1') (Join-Path $Racine 'Modeles\Deploy\sync_startup.ps1') -Force
    Ok 'modeles Cabinet.dotm / Cabinet.xlsm deposes dans Modeles\Deploy'

    # partage reseau (necessite les droits administrateur)
    $estAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $partageExiste = $null -ne (Get-SmbShare -Name $NomPartage -ErrorAction SilentlyContinue)
    if ($partageExiste) { Ok "partage \\$env:COMPUTERNAME\$NomPartage deja present" }
    elseif ($estAdmin) {
        try {
            New-SmbShare -Name $NomPartage -Path $Racine -FullAccess 'Tout le monde' -ErrorAction Stop | Out-Null
            Ok "partage cree : \\$env:COMPUTERNAME\$NomPartage"
        } catch {
            try { New-SmbShare -Name $NomPartage -Path $Racine -FullAccess 'Everyone' -ErrorAction Stop | Out-Null; Ok "partage cree : \\$env:COMPUTERNAME\$NomPartage" }
            catch { Ko "partage non cree ($($_.Exception.Message)) - voir le tuto, etape 'Partager le dossier'" }
        }
    } else {
        Write-Host "  --  pas de droits administrateur : partagez le dossier a la main (tuto) ou relancez PowerShell en administrateur." -ForegroundColor Yellow
        Write-Host "      Chemin a communiquer au poste medecin : \\$env:COMPUTERNAME\$NomPartage"
    }
}
$deploy = Join-Path $Racine 'Modeles\Deploy'
if (-not (Test-Path (Join-Path $deploy 'Cabinet.dotm'))) { Ko "Cabinet.dotm absent de $deploy" }

# ---------------------------------------------------------------- 2. pointeur de racine
Etape 'Enregistrement de la racine sur ce poste'
$dossierApp = Join-Path $env:APPDATA 'CabinetCardio'
New-Item -ItemType Directory -Force -Path $dossierApp | Out-Null
$Racine | Out-File (Join-Path $dossierApp 'chemin.txt') -Encoding ASCII
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0; $word.AutomationSecurity = 1
    $addin = $word.AddIns.Add([string](Join-Path $deploy 'Cabinet.dotm'), $true)
    $arg = [string]$Racine
    $word.Run('EcrireCheminRacine', [ref]$arg)
    try { $addin.Delete() } catch {}
    $word.Quit()
    Ok "chemin.txt ecrit par Word : $Racine"
} catch { Ko "ecriture de chemin.txt par Word impossible : $($_.Exception.Message)" }

# ---------------------------------------------------------------- 3. modeles du poste
Etape 'Installation des modeles Office du poste'
& (Join-Path $PSScriptRoot 'sync_startup.ps1') -Role $Role
$startupDotm = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\Cabinet.dotm'
if (Test-Path $startupDotm) { Ok "Cabinet.dotm dans le demarrage de Word" } else { Ko 'Cabinet.dotm non installe dans STARTUP' }
if ($Role -eq 'Secretaire') {
    $xlsm = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CabinetCardio\Cabinet.xlsm'
    if (Test-Path $xlsm) {
        Ok "Cabinet.xlsm : $xlsm"
        $bureau = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Cabinet Cardio.lnk'
        $sh = New-Object -ComObject WScript.Shell; $lnk = $sh.CreateShortcut($bureau); $lnk.TargetPath = $xlsm; $lnk.Save()
        Ok 'raccourci "Cabinet Cardio" sur le Bureau'
    } else { Ko 'Cabinet.xlsm non installe' }
}

# ---------------------------------------------------------------- 4. mise a jour automatique a l'ouverture de session
Etape 'Tache planifiee de mise a jour (ouverture de session)'
$syncLocal = Join-Path $dossierApp 'sync_startup.ps1'
Copy-Item (Join-Path $PSScriptRoot 'sync_startup.ps1') $syncLocal -Force
$action = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$syncLocal`" -Role $Role"
$res = schtasks /Create /F /TN 'CabinetCardio - mise a jour' /SC ONLOGON /TR $action /RL LIMITED 2>&1
if ($LASTEXITCODE -eq 0) { Ok 'tache "CabinetCardio - mise a jour" creee' } else { Ko "tache planifiee non creee : $res" }

# ---------------------------------------------------------------- 5. raccourcis clavier (medecin)
if ($Role -eq 'Medecin' -and -not $SansRelais) {
    Etape 'Raccourcis clavier Ctrl+Alt+N/C/D/P/G/V dans Normal.dotm'
    try { & (Join-Path $PSScriptRoot 'installer_relais.ps1'); Ok 'relais installe' }
    catch { Ko "relais non installe : $($_.Exception.Message) (verifier l'acces approuve au modele d'objet VBA, tuto etape 3)" }
}

# ---------------------------------------------------------------- 6. cle API (medecin)
if ($Role -eq 'Medecin') {
    Etape 'Cle API Claude'
    $cle = Join-Path $dossierApp 'api.key'
    if (Test-Path $cle) { Ok "cle presente : $cle" }
    else {
        Write-Host "  Collez la cle API (sk-ant-...) puis Entree. Laissez vide pour le faire plus tard." -ForegroundColor Yellow
        $saisie = Read-Host '  Cle API'
        if ($saisie.Trim().Length -gt 20) {
            [IO.File]::WriteAllText($cle, $saisie.Trim(), [Text.Encoding]::ASCII)
            Ok "cle enregistree dans $cle (ce poste uniquement)"
        } else { Write-Host "  --  cle non saisie : creez plus tard le fichier $cle contenant la cle sur une seule ligne." -ForegroundColor Yellow }
    }
}

# ---------------------------------------------------------------- bilan
Etape 'Bilan'
if ($erreurs.Count -eq 0) { Write-Host 'Installation terminee sans erreur.' -ForegroundColor Green }
else { Write-Host "Installation terminee avec $($erreurs.Count) point(s) a corriger :" -ForegroundColor Yellow; $erreurs | ForEach-Object { Write-Host "  - $_" } }
Write-Host ''
Write-Host 'Etapes suivantes : voir TUTO_INSTALLATION (verification, Dragon, ECG, calage CERFA).'
