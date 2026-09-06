# =====================================================================
# verifier_nas.ps1 - Verifie que le logiciel est INTEGRALEMENT present et
# a jour sur le NAS DS224 : depot GitHub clone, code source, donnees de
# configuration, modeles construits. Lecture seule, ne modifie rien.
#
# A executer depuis n'importe quel poste qui voit le NAS (y compris le
# domicile).
# =====================================================================
param(
    [string]$Nas     = '\\DS224\home\claude\claude ai',
    [string]$Depot   = 'https://github.com/Mandagoutolivier/claud-ai-cabinet.git',
    [string]$Branche = 'claude/suivi-dev-logiciel-cabinet-fdjpa9'
)
function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t)    { Write-Host "  OK  $t" -ForegroundColor Green }
function Ko([string]$t)    { Write-Host "  !!  $t" -ForegroundColor Red; $script:problemes += $t }
function Info([string]$t)  { Write-Host "  --  $t" -ForegroundColor Yellow }
$problemes = @()

$git = Join-Path $Nas 'CabinetCardio-Git'
$dev = Join-Path $Nas 'CabinetCardio-Dev'

Etape 'NAS'
if (-not (Test-Path $Nas)) {
    Ko "NAS inaccessible : $Nas"
    Write-Host ''
    Write-Host 'Verification interrompue : impossible de joindre le NAS.' -ForegroundColor Red
    exit 1
}
Ok "NAS accessible : $Nas"

Etape 'Depot GitHub clone sur le NAS'
if (-not (Test-Path (Join-Path $git '.git'))) {
    Ko "pas de clone git dans $git (executez maj_poste.ps1 ou deployer_cabinet.ps1 une fois)"
} else {
    $unc = ($git -replace '\\', '/')
    git config --global --add safe.directory $git 2>$null | Out-Null
    git config --global --add safe.directory $unc 2>$null | Out-Null
    git config --global --add safe.directory "%(prefix)/$unc" 2>$null | Out-Null
    $ici = git -C $git rev-parse HEAD 2>$null
    git -C $git fetch --quiet origin $Branche 2>$null
    $distant = git -C $git rev-parse "origin/$Branche" 2>$null
    if (-not $ici) {
        Ko "impossible de lire le commit courant du clone ($git)"
    } elseif ($ici -eq $distant) {
        Ok "clone a jour avec GitHub ($($ici.Substring(0,8)))"
    } else {
        Ko "clone EN RETARD sur GitHub : local $($ici.Substring(0,8)), origin $($distant.Substring(0,8))"
    }
}

Etape 'Dossier de construction (Src + Donnees)'
foreach ($d in 'Src\Commun', 'Src\Word', 'Src\Excel', 'Src\Forms', 'Donnees\Config', 'Donnees\Modeles') {
    $p = Join-Path $dev $d
    if (Test-Path $p) { Ok "$d present ($((Get-ChildItem $p -Recurse -File).Count) fichiers)" }
    else { Ko "$d ABSENT de $dev" }
}

Etape 'Fichiers de configuration critiques'
foreach ($f in 'config.ini', 'regles_cotation.txt', 'cerfa_positions.txt') {
    $p = Join-Path $dev "Donnees\Config\$f"
    if (Test-Path $p) { Ok "$f present" } else { Ko "$f ABSENT" }
}
foreach ($d in 'demandes', 'profils', 'prompts', 'style', 'gras') {
    $p = Join-Path $dev "Donnees\Config\$d"
    if (Test-Path $p) {
        $n = (Get-ChildItem $p -File -ErrorAction SilentlyContinue).Count
        Ok "Config\$d present ($n fichiers)"
    } else { Ko "Config\$d ABSENT" }
}

Etape 'Modeles Office construits (Modeles\Deploy)'
$deploy = Join-Path $dev 'Donnees\Modeles\Deploy'
foreach ($m in 'Cabinet.dotm', 'Cabinet.xlsm') {
    $p = Join-Path $deploy $m
    if (Test-Path $p) {
        $it = Get-Item $p
        $age = (Get-Date) - $it.LastWriteTime
        if ($age.TotalDays -gt 14) {
            Info "$m present mais date du $($it.LastWriteTime) (plus de 14 jours) - reconstruire avant deploiement"
        } else {
            Ok "$m : $($it.LastWriteTime), $([math]::Round($it.Length/1KB)) Ko"
        }
    } else { Ko "$m ABSENT de $deploy : jamais construit ou construction non publiee" }
}

Etape 'Scripts d installation'
foreach ($f in 'Build\installer_cabinet.ps1', 'Build\build.ps1') {
    $p = Join-Path $dev $f
    if (Test-Path $p) { Ok "$f present" } else { Ko "$f ABSENT" }
}
foreach ($f in 'sync_startup.ps1', 'installer_relais.ps1', 'init_donnees.ps1') {
    $p = Join-Path $dev "Build\$f"
    if (Test-Path $p) { Ok "$f present" } else { Info "$f absent (verifier s'il est necessaire)" }
}

Etape 'Bilan'
if ($problemes.Count -eq 0) {
    Write-Host 'Le logiciel est integralement sauvegarde et deployable depuis le NAS.' -ForegroundColor Green
} else {
    Write-Host "$($problemes.Count) point(s) a corriger avant de considerer le NAS comme deployable :" -ForegroundColor Yellow
    $problemes | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host ''
    Write-Host 'Correction : lancez deployer_cabinet.ps1 (ou maj_poste.ps1) pour rafraichir le NAS depuis GitHub.' -ForegroundColor Yellow
}
