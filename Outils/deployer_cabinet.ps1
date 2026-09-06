# =====================================================================
# deployer_cabinet.ps1 - Deploiement du logiciel au cabinet, DEPUIS LE
# NAS, en un seul point d'execution : le poste du medecin.
#
# 1. Recupere la derniere version depuis GitHub sur le NAS
# 2. Reconstruit Cabinet.dotm / Cabinet.xlsm (Word/Excel ouverts ICI)
# 3. Installe sur CE poste (medecin), racine = partage du secretariat
# 4. Tente de mettre a jour le poste secretariat (RDC) A DISTANCE ;
#    si l'acces distant echoue (pare-feu, WinRM non active...), le
#    paquet est depose sur le partage de RDC avec un lanceur pret a
#    l'emploi que la secretaire n'a qu'a double-cliquer UNE FOIS sur RDC.
#
# Ce dernier point n'est PAS garanti sans WinRM/Planificateur accessibles
# a distance : c'est annonce clairement dans le bilan, ce n'est pas une
# panne du script.
# =====================================================================
param(
    [string]$Nas               = '\\DS224\home\claude\claude ai',
    [string]$Depot             = 'https://github.com/Mandagoutolivier/claud-ai-cabinet.git',
    [string]$Branche           = 'claude/suivi-dev-logiciel-cabinet-fdjpa9',
    [string]$RacineMedecin     = '\\RDC\CabinetCardio',
    [string]$PosteSecretariat  = 'RDC',
    [string]$RacineSecretariat = 'C:\CabinetCardio',
    [switch]$SansConstruction,           # reutilise les modeles deja construits
    [switch]$SansSecretariat             # ne touche pas au poste RDC
)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t)    { Write-Host "  OK  $t" -ForegroundColor Green }
function Ko([string]$t)    { Write-Host "  !!  $t" -ForegroundColor Red; $script:erreurs += $t }
function Info([string]$t)  { Write-Host "  --  $t" -ForegroundColor Yellow }
$erreurs = @()

# ---------------------------------------------------------------------
# Complete un config.ini de poste avec les cles ABSENTES du config.ini de
# reference (celui du paquet), sans jamais modifier une valeur presente.
# Les commentaires qui precedent une cle dans la reference sont recopies.
# Une section absente est ajoutee en fin de fichier.
# ---------------------------------------------------------------------
function Completer-ConfigIni([string]$reference, [string]$cible) {
    if (-not (Test-Path $reference)) { return }
    if (-not (Test-Path $cible)) { Copy-Item $reference $cible; Ok "config.ini cree : $cible"; return }

    function Lire-Ini([string]$chemin) {
        $sections = [ordered]@{}; $section = ''; $commentaires = @()
        foreach ($ligne in [IO.File]::ReadAllLines($chemin, [Text.Encoding]::UTF8)) {
            $t = $ligne.Trim()
            if ($t -match '^\[(.+)\]$') { $section = $matches[1]; if (-not $sections.Contains($section)) { $sections[$section] = [ordered]@{} }; $commentaires = @(); continue }
            if ($t -eq '' -or $t.StartsWith(';') -or $t.StartsWith('#')) { if ($t -ne '') { $commentaires += $ligne } else { $commentaires = @() }; continue }
            if ($t -match '^([^=]+?)\s*=(.*)$') {
                if ($section -eq '') { continue }
                $sections[$section][$matches[1].Trim()] = @{ ligne = $ligne; commentaires = $commentaires }
                $commentaires = @()
            }
        }
        return $sections
    }

    $ref = Lire-Ini $reference
    $act = Lire-Ini $cible
    $lignes = [Collections.Generic.List[string]]::new([IO.File]::ReadAllLines($cible, [Text.Encoding]::UTF8))
    $ajouts = 0

    foreach ($section in $ref.Keys) {
        $manquantes = @()
        foreach ($cle in $ref[$section].Keys) {
            if (-not ($act.Contains($section) -and $act[$section].Contains($cle))) { $manquantes += $cle }
        }
        if ($manquantes.Count -eq 0) { continue }

        # position d'insertion : fin de la section existante, sinon fin du fichier
        $idx = -1
        for ($i = 0; $i -lt $lignes.Count; $i++) { if ($lignes[$i].Trim() -eq "[$section]") { $idx = $i; break } }
        $bloc = @()
        if ($idx -lt 0) {
            $bloc += ''; $bloc += "[$section]"
            $insertion = $lignes.Count
        } else {
            $insertion = $lignes.Count
            for ($i = $idx + 1; $i -lt $lignes.Count; $i++) { if ($lignes[$i].Trim() -match '^\[.+\]$') { $insertion = $i; break } }
            # remonter au-dessus des lignes vides qui separent les sections
            while ($insertion -gt $idx + 1 -and $lignes[$insertion - 1].Trim() -eq '') { $insertion-- }
        }
        foreach ($cle in $manquantes) {
            $bloc += $ref[$section][$cle].commentaires
            $bloc += $ref[$section][$cle].ligne
            $ajouts++
        }
        $lignes.InsertRange($insertion, [string[]]$bloc)
    }

    if ($ajouts -gt 0) {
        Copy-Item $cible "$cible.avant_$(Get-Date -Format yyyyMMdd-HHmmss).bak" -Force
        [IO.File]::WriteAllLines($cible, $lignes, (New-Object Text.UTF8Encoding($false)))
        Ok "config.ini complete : $ajouts cle(s) ajoutee(s) (valeurs existantes intactes, copie .bak conservee)"
    } else {
        Ok 'config.ini deja complet'
    }
}

$git   = Join-Path $Nas 'CabinetCardio-Git'
$dev   = Join-Path $Nas 'CabinetCardio-Dev'
$build = Join-Path $dev 'Build'
$deployNas = Join-Path $dev 'Donnees\Modeles\Deploy'

# ------------------------------------------------------------ 0. controles
Etape 'Controles'
if (Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue) {
    Write-Host 'Fermez COMPLETEMENT Word et Excel sur ce poste puis relancez.' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $Nas)) { throw "NAS inaccessible : $Nas (le DS224 est-il allume et le lecteur connecte au cabinet ?)" }
Ok "NAS accessible : $Nas"
if (-not (Get-Command git.exe -ErrorAction SilentlyContinue)) {
    throw "git introuvable sur ce poste. Installez Git pour Windows puis relancez."
}
Ok 'git present'
if (-not (Test-Path $RacineMedecin)) {
    throw "Racine du secretariat inaccessible : $RacineMedecin (le PC $PosteSecretariat est-il allume et le dossier partage ?)"
}
Ok "secretariat joignable : $RacineMedecin"

function Assurer-SafeDirectory([string]$chemin) {
    $unc = $chemin -replace '\\', '/'
    $formes = @($chemin, $unc, "%(prefix)/$unc")
    $deja = @(git config --global --get-all safe.directory 2>$null)
    foreach ($f in $formes) {
        if ($deja -notcontains $f) { git config --global --add safe.directory $f | Out-Null }
    }
    $global:LASTEXITCODE = 0
}
Assurer-SafeDirectory $git

# ------------------------------------------------------------ 1. GitHub -> NAS
Etape "Recuperation de la version GitHub (branche $Branche)"
if (Test-Path (Join-Path $git '.git')) {
    git -C $git remote set-url origin $Depot
    git -C $git fetch --prune origin $Branche
    git -C $git checkout -B $Branche "origin/$Branche"
    git -C $git reset --hard "origin/$Branche"
} else {
    if (Test-Path $git) { Rename-Item $git "$git.old_$(Get-Date -Format yyyyMMdd-HHmmss)" }
    git clone --branch $Branche --single-branch $Depot $git
}
if ($LASTEXITCODE -ne 0) { throw 'Recuperation GitHub echouee (identifiants GitHub ? reseau internet du cabinet ?).' }
$version = (git -C $git log -1 --format='%h %ad %s' --date=short)
Ok "version recuperee : $version"

# ------------------------------------------------------------ 2. NAS : dossier de construction
Etape "Mise a jour du dossier de construction : $dev"
New-Item -ItemType Directory -Force -Path $dev | Out-Null
robocopy (Join-Path $git 'Src') (Join-Path $dev 'Src') /MIR /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie de Src vers le NAS echouee (code $LASTEXITCODE)." }
Ok 'Src copie (miroir)'
robocopy (Join-Path $git 'Donnees') (Join-Path $dev 'Donnees') /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 `
    /XD 'Patients' 'Actes' 'Echange' 'Sauvegardes' 'Logs' | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie de Donnees vers le NAS echouee (code $LASTEXITCODE)." }
Ok 'Donnees copie (configuration et modeles ; bases existantes conservees)'
foreach ($f in 'installer_cabinet.ps1', 'README.md', 'LETTRES_DERIVEES.md', 'OPTIMISATIONS_20260905.md', 'RECETTE_AUDIT.md') {
    $s = Join-Path $git $f
    if (Test-Path $s) { Copy-Item $s (Join-Path $build $f) -Force -ErrorAction SilentlyContinue }
}
$global:LASTEXITCODE = 0

# ------------------------------------------------------------ 3. construction
$sortie = Join-Path $build 'out'
$avant = @{}
foreach ($m in 'Cabinet.dotm', 'Cabinet.xlsm') {
    $p = Join-Path $sortie $m
    $avant[$m] = if (Test-Path $p) { (Get-Item $p).LastWriteTime } else { [datetime]::MinValue }
}
if (-not $SansConstruction) {
    Etape 'Construction de Cabinet.dotm et Cabinet.xlsm (sur ce poste)'
    $script = Join-Path $build 'build.ps1'
    if (-not (Test-Path $script)) { throw "build.ps1 introuvable : $script" }
    Info 'Word et Excel vont s ouvrir automatiquement : ne touchez a rien pendant la construction.'
    & $script
    if ($LASTEXITCODE -ne 0) { throw "Construction echouee (acces approuve au modele d'objet VBA active dans Word ET Excel ?)." }
}
New-Item -ItemType Directory -Force -Path $deployNas | Out-Null
foreach ($m in 'Cabinet.dotm', 'Cabinet.xlsm') {
    $p = Join-Path $sortie $m
    if (-not (Test-Path $p)) { throw "$m absent de $sortie : la construction n'a pas abouti." }
    $date = (Get-Item $p).LastWriteTime
    if (-not $SansConstruction -and $date -le $avant[$m]) {
        throw ("$m N'A PAS ETE RECONSTRUIT (toujours du $date).`n" +
               "Lancez `"$script`" a la main et lisez son message d'erreur.")
    }
    Copy-Item $p (Join-Path $deployNas $m) -Force
    Ok "$m construit le $date, publie dans Modeles\Deploy"
}

# ------------------------------------------------------------ 4. paquet local + installation MEDECIN
Etape 'Preparation du paquet d installation sur ce poste'
$paquet = Join-Path $env:LOCALAPPDATA 'CabinetCardio-Install'
if (Test-Path $paquet) { Remove-Item $paquet -Recurse -Force }
robocopy $dev $paquet /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie du paquet en local echouee (code $LASTEXITCODE)." }
$global:LASTEXITCODE = 0
Ok "paquet : $paquet"

$installeur = Join-Path $paquet 'Build\installer_cabinet.ps1'
if (-not (Test-Path $installeur)) { throw "installer_cabinet.ps1 introuvable : $installeur" }
Etape "Installation sur ce poste (Medecin, racine $RacineMedecin)"
& $installeur -Role Medecin -Racine $RacineMedecin

# la racine est partagee avec le secretariat : un seul config.ini a completer
Etape 'Configuration : cles nouvelles'
Completer-ConfigIni (Join-Path $paquet 'Donnees\Config\config.ini') (Join-Path $RacineMedecin 'Config\config.ini')

# ------------------------------------------------------------ 5. poste SECRETARIAT (RDC), best effort
$secOk = $false
if ($SansSecretariat) {
    Etape 'Poste secretariat'
    Info "ignore (-SansSecretariat) : mettez-le a jour separement avec maj_poste.ps1 -Role Secretaire"
} else {
    Etape "Poste secretariat ($PosteSecretariat) : depot du paquet"
    $partageSecretariat = "\\$PosteSecretariat\CabinetCardio"
    if (-not (Test-Path $partageSecretariat)) {
        Ko "partage $partageSecretariat inaccessible : impossible de deposer le paquet"
    } else {
        $depot = Join-Path $partageSecretariat '_Installation'
        try {
            robocopy $paquet $depot /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
            if ($LASTEXITCODE -ge 8) { throw "code $LASTEXITCODE" }
            $global:LASTEXITCODE = 0
            @"
@echo off
title Installation Cabinet Cardio - poste secretariat
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\installer_cabinet.ps1" -Role Secretaire -Racine "$RacineSecretariat"
echo.
pause
"@ | Out-File (Join-Path $depot 'installer_secretariat.cmd') -Encoding ASCII
            Ok "paquet depose sur $depot, lanceur installer_secretariat.cmd pret"
        } catch {
            Ko "depot du paquet sur $partageSecretariat echoue : $($_.Exception.Message)"
        }

        # tentative d'execution a distance (best effort : necessite WinRM ou
        # le Planificateur de taches accessible a distance, non garanti)
        Etape "Poste secretariat ($PosteSecretariat) : tentative d'installation a distance"
        try {
            $ancien = $ErrorActionPreference; $ErrorActionPreference = 'Stop'
            $accessible = Test-WSMan -ComputerName $PosteSecretariat -ErrorAction Stop
            $ErrorActionPreference = $ancien
            $cred = Get-Credential -Message "Compte administrateur de $PosteSecretariat (ex: $PosteSecretariat\PATRICIA)"
            Invoke-Command -ComputerName $PosteSecretariat -Credential $cred -ScriptBlock {
                param($racine)
                & "C:\CabinetCardio\_Installation\Build\installer_cabinet.ps1" -Role Secretaire -Racine $racine
            } -ArgumentList $RacineSecretariat
            Ok "installation executee a distance sur $PosteSecretariat"
            $secOk = $true
        } catch {
            $ErrorActionPreference = 'Stop'
            Info "installation a distance impossible ($($_.Exception.Message))"
            Info "ACTION MANUELLE REQUISE UNE FOIS : sur le PC $PosteSecretariat, ouvrir"
            Info "  \\$PosteSecretariat\CabinetCardio\_Installation\installer_secretariat.cmd"
            Info "et double-cliquer (fermer Word/Excel sur ce poste avant)."
        }
    }
}

# ------------------------------------------------------------ bilan
Etape 'Bilan'
Ok "version deployee : $version"
Write-Host "  Poste medecin (ce PC)     : installe, racine $RacineMedecin"
if ($secOk) {
    Write-Host "  Poste secretariat ($PosteSecretariat) : installe automatiquement" -ForegroundColor Green
} elseif ($SansSecretariat) {
    Write-Host "  Poste secretariat ($PosteSecretariat) : NON TRAITE (option -SansSecretariat)" -ForegroundColor Yellow
} else {
    Write-Host "  Poste secretariat ($PosteSecretariat) : A FINALISER A LA MAIN sur ce poste (voir ci-dessus)" -ForegroundColor Yellow
}
if ($erreurs.Count -gt 0) {
    Write-Host ''
    Write-Host "$($erreurs.Count) point(s) a corriger :" -ForegroundColor Yellow
    $erreurs | ForEach-Object { Write-Host "  - $_" }
}
Write-Host ''
Write-Host 'Ouvrez Word sur ce poste : le ruban "Cabinet" doit apparaitre.' -ForegroundColor Green
