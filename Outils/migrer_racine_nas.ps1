# =====================================================================
# migrer_racine_nas.ps1 - Deplace la racine des donnees du cabinet
# (bases patients / agenda / journal, dossiers patients, config, echanges)
# du PC secretariat vers un dossier partage du NAS DS224, puis enregistre
# ce poste sur la nouvelle racine.
#
# A lancer :
#   1. sur ACCUEIL (Word/Excel fermes sur les DEUX postes) :
#        powershell -NoProfile -ExecutionPolicy Bypass -File .\migrer_racine_nas.ps1 -Role Secretaire
#      -> copie integrale C:\CabinetCardio -> \\DS224\CabinetCardio, verification,
#         puis l'ancien dossier local est renomme CabinetCardio_ANCIEN_<date>
#         (rien n'est supprime) et le poste pointe sur le NAS ;
#   2. sur AX8_Max :
#        powershell -NoProfile -ExecutionPolicy Bypass -File .\migrer_racine_nas.ps1 -Role Medecin
#      -> le poste pointe sur le NAS (aucune copie : deja faite).
# Prerequis : dossier partage "CabinetCardio" cree dans DSM (Panneau de
# configuration > Dossier partage), lecture/ecriture pour les comptes des
# deux postes, identifiants enregistres dans Windows (Gestionnaire
# d'identification) pour ne pas etre redemandes.
# =====================================================================
param(
    [ValidateSet('Secretaire', 'Medecin', 'Tous')][string]$Role = 'Secretaire',
    [string]$RacineNas = '\\DS224\CabinetCardio',
    [string]$AncienneRacine = ''
)
$ErrorActionPreference = 'Stop'
trap { Write-Host ''; Write-Host "ARRET : $($_.Exception.Message)" -ForegroundColor Red; Read-Host 'Appuyez sur Entree pour fermer cette fenetre'; exit 1 }
function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t)    { Write-Host "  OK  $t" -ForegroundColor Green }
function Info([string]$t)  { Write-Host "  --  $t" -ForegroundColor Yellow }

$dossierApp = Join-Path $env:APPDATA 'CabinetCardio'
$cheminTxt = Join-Path $dossierApp 'chemin.txt'
if ($AncienneRacine -eq '' -and (Test-Path $cheminTxt)) { $AncienneRacine = (Get-Content $cheminTxt -TotalCount 1).Trim() }
if ($AncienneRacine -eq '') { $AncienneRacine = 'C:\CabinetCardio' }

Etape 'Controles'
if (Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue) { throw 'Fermez Word et Excel sur ce poste (et sur l autre poste) avant de migrer.' }
$parent = Split-Path $RacineNas -Parent
if (-not (Test-Path $parent)) { throw "NAS inaccessible : $parent - creez le dossier partage dans DSM et verifiez les identifiants." }
Ok "NAS accessible : $parent"
if ($AncienneRacine -ieq $RacineNas) { throw "Ce poste pointe deja sur $RacineNas : rien a migrer." }
Ok "racine actuelle de ce poste : $AncienneRacine"

if ($Role -ne 'Medecin') {
    if (-not (Test-Path (Join-Path $AncienneRacine 'Base\Patients.xlsx'))) {
        throw "Base\Patients.xlsx introuvable dans $AncienneRacine : ce script se lance sur le poste qui heberge les donnees (ACCUEIL)."
    }
    if (Test-Path (Join-Path $RacineNas 'Base\Patients.xlsx')) {
        Info "$RacineNas contient deja une base patients : la copie est ADDITIVE (aucun fichier existant du NAS n'est ecrase)."
        if ((Read-Host 'Continuer ? (O/N)') -notmatch '^[oO]') { exit 0 }
    }
    Etape "Copie $AncienneRacine -> $RacineNas"
    New-Item -ItemType Directory -Force -Path $RacineNas | Out-Null
    # /E tout, /XO n'ecrase que par plus recent, jamais de suppression, verrous exclus
    robocopy $AncienneRacine $RacineNas /E /XO /XD locks _Installation /XF *.lock /R:3 /W:2 /NP /NFL /NDL
    if ($LASTEXITCODE -ge 8) { throw "robocopy a echoue (code $LASTEXITCODE) : rien n'a ete bascule." }
    $global:LASTEXITCODE = 0
    Etape 'Verification'
    $manquants = 0
    foreach ($f in Get-ChildItem $AncienneRacine -Recurse -File | Where-Object { $_.FullName -notmatch '\\locks\\|\\_Installation\\|\.lock$' }) {
        $rel = $f.FullName.Substring($AncienneRacine.Length + 1)
        $d = Join-Path $RacineNas $rel
        if (-not (Test-Path $d) -or (Get-Item $d).Length -ne $f.Length) { $manquants++; Write-Host "  !!  $rel" -ForegroundColor Red }
    }
    if ($manquants -gt 0) { throw "$manquants fichier(s) different(s) ou absents sur le NAS : l'ancienne racine est conservee, rien n'a ete bascule." }
    Ok "tous les fichiers sont presents sur le NAS a l'identique"
    New-Item -ItemType Directory -Force -Path (Join-Path $RacineNas 'Base\locks') | Out-Null

    if (-not $AncienneRacine.StartsWith('\\')) {
        Etape 'Neutralisation de l ancienne racine locale'
        $ancienNom = "$AncienneRacine`_ANCIEN_$(Get-Date -Format yyyyMMdd)"
        try {
            Rename-Item $AncienneRacine $ancienNom -ErrorAction Stop
            Ok "renommee en $ancienNom (a supprimer apres quelques semaines de fonctionnement)"
        } catch {
            Info "renommage impossible ($($_.Exception.Message)) : partage encore ouvert ? Renommez-la a la main plus tard."
        }
        $partage = Get-SmbShare -Name 'CabinetCardio' -ErrorAction SilentlyContinue
        if ($partage) {
            try { Remove-SmbShare -Name 'CabinetCardio' -Force -ErrorAction Stop; Ok 'partage \\' + $env:COMPUTERNAME + '\CabinetCardio supprime' }
            catch { Info "partage CabinetCardio non supprime (droits administrateur requis) : a faire a la main pour eviter tout usage de l'ancienne copie." }
        }
    }
}

Etape "Enregistrement de ce poste sur $RacineNas"
New-Item -ItemType Directory -Force -Path $dossierApp | Out-Null
$RacineNas | Out-File $cheminTxt -Encoding ASCII
Ok "chemin.txt -> $RacineNas"
$sync = Join-Path $dossierApp 'sync_startup.ps1'
if (Test-Path $sync) {
    $roleSync = if ($Role -eq 'Medecin') { 'Medecin' } else { 'Tous' }
    & $sync -Role $roleSync
    Ok 'modeles Office resynchronises depuis la nouvelle racine'
}

Etape 'Bilan'
Write-Host "  Ce poste travaille desormais sur $RacineNas" -ForegroundColor Green
if ($Role -ne 'Medecin') {
    Write-Host '  A FAIRE sur le poste medecin (AX8_Max) : lancer ce meme script avec -Role Medecin,' -ForegroundColor Yellow
    Write-Host '  puis "Deployer le cabinet" (il detecte la racine NAS et y depose le paquet d installation).' -ForegroundColor Yellow
}
Write-Host '  Sauvegarde : activer sur le DS224 une tache Hyper Backup du dossier partage CabinetCardio' -ForegroundColor Yellow
Write-Host '  vers le NAS du domicile (quotidienne, versions conservees 30 jours).' -ForegroundColor Yellow
Read-Host 'Appuyez sur Entree pour fermer cette fenetre'
