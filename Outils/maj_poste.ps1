# =====================================================================
# maj_poste.ps1 - Recupere la version GitHub du logiciel, la depose sur
# le NAS DS224, reconstruit Cabinet.dotm / Cabinet.xlsm et installe le
# resultat sur CE poste.
#
# Utilisation courante (poste du domicile, autonome) :
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\maj_poste.ps1
#
# Poste medecin relie au secretariat :
#   .\maj_poste.ps1 -Racine "\\RDC\CabinetCardio" -Role Medecin
#
# Le script ne detruit jamais une base de donnees existante : les
# fichiers deja presents dans la racine sont conserves.
# =====================================================================
param(
    [string]$Nas      = '\\DS224\home\claude\claude ai',
    [string]$Depot    = 'https://github.com/Mandagoutolivier/claud-ai-cabinet.git',
    [string]$Branche  = 'claude/suivi-dev-logiciel-cabinet-fdjpa9',
    [string]$Racine   = 'C:\CabinetCardio',
    [ValidateSet('Auto', 'Medecin', 'Secretaire')][string]$Role = 'Auto',
    # Base patients d'amorcage, utilisee UNIQUEMENT si la racine n'en a pas
    # encore. Les bases nominatives ne sont pas dans GitHub : elles viennent
    # du paquet d'installation du NAS ou du poste secretariat.
    [string]$BaseInitiale = '\\DS224\home\claude\install\Donnees\Base\Patients.xlsx',
    [switch]$SansConstruction     # reutilise les modeles deja construits
)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t)    { Write-Host "  OK  $t" -ForegroundColor Green }
function Info([string]$t)  { Write-Host "  --  $t" -ForegroundColor Yellow }

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

$git   = Join-Path $Nas 'CabinetCardio-Git'     # clone GitHub (jamais modifie a la main)
$dev   = Join-Path $Nas 'CabinetCardio-Dev'     # dossier de construction
$build = Join-Path $dev 'Build'

# ------------------------------------------------------------ 0. controles
Etape 'Controles'
if (Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue) {
    Write-Host 'Fermez COMPLETEMENT Word et Excel (fenetres reduites comprises) puis relancez.' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $Nas)) { throw "NAS inaccessible : $Nas (le DS224 est-il allume et le lecteur connecte ?)" }
Ok "NAS accessible : $Nas"
if (-not (Get-Command git.exe -ErrorAction SilentlyContinue)) {
    throw "git introuvable. Installez Git pour Windows (https://git-scm.com/download/win) puis relancez."
}
Ok 'git present'

# Un depot git pose sur un partage reseau appartient a l'utilisateur du NAS,
# pas a celui de Windows : git refuse alors d'y toucher ("dubious ownership").
# On declare le dossier comme sur, sous les deux ecritures que git accepte
# pour un chemin UNC. Operation idempotente et sans effet sur les autres depots.
function Assurer-SafeDirectory([string]$chemin) {
    $unc = $chemin -replace '\\', '/'          # \\DS224\... -> //DS224/...
    $formes = @($chemin, $unc, "%(prefix)/$unc")
    $deja = @(git config --global --get-all safe.directory 2>$null)
    foreach ($f in $formes) {
        if ($deja -notcontains $f) { git config --global --add safe.directory $f | Out-Null }
    }
    $global:LASTEXITCODE = 0
}
Assurer-SafeDirectory $git
Ok 'depot du NAS declare comme sur pour git (safe.directory)'

# ------------------------------------------------------------ 1. recuperation GitHub -> NAS
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
if ($LASTEXITCODE -ne 0) { throw 'Recuperation GitHub echouee (identifiants GitHub ? reseau ?).' }
$version = (git -C $git log -1 --format='%h %ad %s' --date=short)
Ok "version recuperee : $version"

# ------------------------------------------------------------ 2. NAS : dossier de construction
Etape "Mise a jour du dossier de construction : $dev"
New-Item -ItemType Directory -Force -Path $dev | Out-Null
# ATTENTION : Src est MIROIR (le code vient de GitHub, rien d autre n a a y
# vivre), mais Donnees est copie SANS /MIR et SANS toucher aux bases : les
# fichiers nominatifs (Patients.xlsx, Journal_*.xlsx, Agenda_*.xlsx, dossiers
# patients) ne sont PAS dans GitHub, ils ne doivent jamais etre effaces du NAS.
# robocopy (et non Copy-Item) : gere les sous-dossiers, les accents et les
# chemins UNC longs, et ne bute pas sur "le conteneur ne peut pas etre copie".
robocopy (Join-Path $git 'Src') (Join-Path $dev 'Src') /MIR /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie de Src vers le NAS echouee (code $LASTEXITCODE)." }
Ok 'Src copie (miroir)'

# Donnees : on ajoute et on met a jour, on ne supprime jamais.
# /XD : dossiers de donnees vivantes exclus de la mise a jour.
robocopy (Join-Path $git 'Donnees') (Join-Path $dev 'Donnees') /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 `
    /XD 'Patients' 'Actes' 'Echange' 'Sauvegardes' 'Logs' | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie de Donnees vers le NAS echouee (code $LASTEXITCODE)." }
Ok 'Donnees copie (configuration et modeles ; bases existantes conservees)'

$basePat = Join-Path $dev 'Donnees\Base\Patients.xlsx'
if (Test-Path $basePat) { Ok "base patients du NAS conservee : $basePat" }
else { Info "pas de Patients.xlsx dans $dev\Donnees\Base : le logiciel creera une base vide a la premiere utilisation" }

foreach ($f in 'installer_cabinet.ps1', 'README.md', 'LETTRES_DERIVEES.md', 'OPTIMISATIONS_20260905.md') {
    $s = Join-Path $git $f
    if (Test-Path $s) { Copy-Item $s (Join-Path $build $f) -Force -ErrorAction SilentlyContinue }
}
if (Test-Path (Join-Path $dev 'Donnees\Config\profils')) {
    $n = (Get-ChildItem (Join-Path $dev 'Donnees\Config\profils') -Filter *.ini).Count
    Ok "$n profils de demande d'examen presents"
} else { Info 'Donnees\Config\profils absent : verifiez la copie ci-dessus' }
$LASTEXITCODE = 0

# ------------------------------------------------------------ 3. construction des modeles
$deployNas = Join-Path $dev 'Donnees\Modeles\Deploy'
# build.ps1 ecrit les modeles dans Build\out : c'est LA sortie de reference.
# Donnees\Modeles\Deploy n'est qu'une copie de distribution, qu'il faut
# rafraichir apres chaque construction (sinon on reinstalle l'ancien modele).
$sortie = Join-Path $build 'out'
$avant = @{}
foreach ($m in 'Cabinet.dotm', 'Cabinet.xlsm') {
    $p = Join-Path $sortie $m
    $avant[$m] = if (Test-Path $p) { (Get-Item $p).LastWriteTime } else { [datetime]::MinValue }
}
if (-not $SansConstruction) {
    Etape 'Construction de Cabinet.dotm et Cabinet.xlsm'
    $script = Join-Path $build 'build.ps1'
    if (-not (Test-Path $script)) { throw "build.ps1 introuvable : $script" }
    Info 'Word et Excel vont s ouvrir automatiquement : ne touchez a rien pendant la construction.'
    & $script
    if ($LASTEXITCODE -ne 0) { throw "Construction echouee (acces approuve au modele d'objet VBA activee dans Word ET Excel ?)." }
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

# ------------------------------------------------------------ 4. paquet d'installation local
Etape 'Preparation du paquet d installation sur ce poste'
$paquet = Join-Path $env:LOCALAPPDATA 'CabinetCardio-Install'
if (Test-Path $paquet) { Remove-Item $paquet -Recurse -Force }
robocopy $dev $paquet /E /NFL /NDL /NJH /NJS /NP /R:2 /W:2 | Out-Null
if ($LASTEXITCODE -ge 8) { throw "Copie du paquet en local echouee (code $LASTEXITCODE)." }
$LASTEXITCODE = 0
Ok "paquet : $paquet"

# ------------------------------------------------------------ 5. installation
$installeur = Join-Path $paquet 'Build\installer_cabinet.ps1'
if (-not (Test-Path $installeur)) { throw "installer_cabinet.ps1 introuvable : $installeur" }

$racineDistante = $Racine.StartsWith('\\')
if ($Role -eq 'Auto') { $Role = if ($racineDistante) { 'Medecin' } else { 'Secretaire' } }

$basePoste = Join-Path $Racine 'Base\Patients.xlsx'
if (-not $racineDistante -and -not (Test-Path $basePoste) -and (Test-Path $BaseInitiale)) {
    Etape 'Base patients initiale'
    New-Item -ItemType Directory -Force -Path (Split-Path $basePoste -Parent) | Out-Null
    Copy-Item $BaseInitiale $basePoste
    $src = Get-Item $BaseInitiale
    Ok "base copiee depuis $BaseInitiale ($([math]::Round($src.Length/1KB)) Ko, du $($src.LastWriteTime))"
    Info 'ATTENTION : cette copie est une PHOTO. Ce que vous saisirez ici ne remontera pas au cabinet.'
}

if (-not $racineDistante -and -not (Test-Path (Join-Path $Racine 'Base'))) {
    # poste autonome (domicile) : on cree d abord la racine locale et ses bases
    Etape "Creation de la racine locale $Racine"
    & $installeur -Role Secretaire -Racine $Racine
}
Etape "Installation sur ce poste (role $Role, racine $Racine)"
& $installeur -Role $Role -Racine $Racine

# nouvelles cles de configuration apportees par cette version : ajoutees au
# config.ini du poste sans toucher aux reglages deja saisis
Etape 'Configuration : cles nouvelles'
Completer-ConfigIni (Join-Path $paquet 'Donnees\Config\config.ini') (Join-Path $Racine 'Config\config.ini')

# ------------------------------------------------------------ bilan
Etape 'Bilan'
Ok "version installee : $version"
Write-Host "  Donnees   : $Racine"
Write-Host "  Modeles   : $(Join-Path $Racine 'Modeles\Deploy')"
Write-Host ''
Write-Host 'Ouvrez Word : le ruban "Cabinet" doit apparaitre (Ctrl+Alt+N nouveau courrier).' -ForegroundColor Green
