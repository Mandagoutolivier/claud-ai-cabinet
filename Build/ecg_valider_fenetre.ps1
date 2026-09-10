# =====================================================================
# ecg_valider_fenetre.ps1 - Veilleur du poste MEDECIN (AX8_Max).
# Quand la fenetre patient de Resting12Lead apparait apres l'import du
# GDT, la date de naissance affichee n'est prise en compte que si le
# champ est "valide" (tabulation). Ce script tourne en arriere-plan,
# detecte l'apparition d'une NOUVELLE fenetre dont le titre contient
# [ECG] FenetreTitre, l'active et lui envoie [ECG] TouchesValidation
# (syntaxe SendKeys : {TAB}, {ENTER}...), une seule fois par fenetre.
#
# Reglages dans Config\config.ini ([ECG]) :
#   FenetreTitre=Resting12Lead      (vide = veilleur desactive)
#   TouchesValidation={TAB}{TAB}
#   DelaiFenetreMs=1500             (attente apres apparition)
# Lance a l'ouverture de session par installer_cabinet.ps1 (role Medecin).
# =====================================================================
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName Microsoft.VisualBasic

function LireIni([string]$fichier, [string]$section, [string]$cle, [string]$defaut) {
    if (-not (Test-Path $fichier)) { return $defaut }
    $dans = $false
    foreach ($l in Get-Content $fichier -Encoding UTF8) {
        $t = $l.Trim()
        if ($t -match '^\[(.+)\]$') { $dans = ($Matches[1] -ieq $section); continue }
        if ($dans -and $t -match '^([^;#=]+)=(.*)$' -and $Matches[1].Trim() -ieq $cle) { return $Matches[2].Trim() }
    }
    return $defaut
}

$cheminTxt = Join-Path $env:APPDATA 'CabinetCardio\chemin.txt'
if (-not (Test-Path $cheminTxt)) { exit 0 }
$racine = (Get-Content $cheminTxt -TotalCount 1).Trim()
$ini = Join-Path $racine 'Config\config.ini'
$wsh = New-Object -ComObject WScript.Shell
$vus = @{}
while ($true) {
    # relu a chaque tour : un changement de reglage est pris en compte sans relance
    $titre   = LireIni $ini 'ECG' 'FenetreTitre' ''
    $touches = LireIni $ini 'ECG' 'TouchesValidation' '{TAB}{TAB}'
    $delai   = [int](LireIni $ini 'ECG' 'DelaiFenetreMs' '1500')
    if ($titre.Length -gt 0 -and $touches.Length -gt 0) {
        foreach ($p in Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -like "*$titre*" }) {
            $cle = "$($p.Id)-$($p.MainWindowHandle)-$($p.MainWindowTitle)"
            if ($vus.ContainsKey($cle)) { continue }
            $vus[$cle] = Get-Date
            Start-Sleep -Milliseconds $delai
            try {
                [Microsoft.VisualBasic.Interaction]::AppActivate($p.Id)
                Start-Sleep -Milliseconds 200
                $wsh.SendKeys($touches)
            } catch {}
        }
        # oubli des fenetres disparues (pour re-valider une prochaine ouverture)
        foreach ($k in @($vus.Keys)) {
            $pid0 = [int]($k -split '-')[0]
            if (-not (Get-Process -Id $pid0 -ErrorAction SilentlyContinue)) { $vus.Remove($k) }
        }
    }
    Start-Sleep -Milliseconds 700
}
