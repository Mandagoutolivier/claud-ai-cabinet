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
$journal = Join-Path $racine 'Logs\ecg_sql.log'
$wsh = New-Object -ComObject WScript.Shell
$vus = @{}

function Journal([string]$t) {
    try { Add-Content -Path $journal -Value ("{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $t) -Encoding UTF8 } catch {}
}

# --- pont GDT -> SQL (boite aux lettres "patinfo" lue par Resting12Lead) ---
# A chaque nouvel IMPORT.GDT ecrit dans [ECG] DossierGdt (par "Arrive" du
# secretariat ou Ctrl+Alt+G du medecin), la ligne du patient est (re)ecrite
# dans la table patinfo de la base [ECG] SqlBase sur l'instance [ECG]
# SqlInstance (SQL Server Express local). Le sexe, absent de l'import GDT
# de Resting12Lead, passe ainsi par la connexion "Systeme Info DMS".
function LireGdt([string]$fichier) {
    $d = @{}
    foreach ($l in [IO.File]::ReadAllLines($fichier, [Text.Encoding]::GetEncoding(1252))) {
        if ($l.Length -ge 7) { $d[$l.Substring(3, 4)] = $l.Substring(7) }
    }
    return $d
}

function EcrireSql([hashtable]$g) {
    $instance = LireIni $ini 'ECG' 'SqlInstance' '.\SQLEXPRESS'
    $base     = LireIni $ini 'ECG' 'SqlBase' 'ecgcenter'
    $user     = LireIni $ini 'ECG' 'SqlUtilisateur' 'DMSNIS'
    $mdp      = LireIni $ini 'ECG' 'SqlMotDePasse' 'dmsdms'
    $table    = LireIni $ini 'ECG' 'SqlTable' 'patinfo'
    $sexeM    = LireIni $ini 'ECG' 'SqlSexeM' 'M'
    $sexeF    = LireIni $ini 'ECG' 'SqlSexeF' 'F'
    if (-not $g.ContainsKey('3000')) { throw 'GDT sans identifiant patient (3000)' }
    $sexe = ''
    switch ($g['3110']) { '1' { $sexe = $sexeM } '2' { $sexe = $sexeF } 'M' { $sexe = $sexeM } 'W' { $sexe = $sexeF } 'F' { $sexe = $sexeF } }
    $j = $null; $m = $null; $a = $null; $age = $null
    if ($g['3103'] -match '^(\d{2})(\d{2})(\d{4})$') {
        $j = [int]$Matches[1]; $m = [int]$Matches[2]; $a = [int]$Matches[3]
        $ddn = Get-Date -Year $a -Month $m -Day $j
        $age = [int][math]::Floor(((Get-Date) - $ddn).TotalDays / 365.25)
    }
    $cs = if ($user.Length -gt 0) { "Server=$instance;Database=$base;User ID=$user;Password=$mdp;Connect Timeout=5" }
          else { "Server=$instance;Database=$base;Integrated Security=True;Connect Timeout=5" }
    $cn = New-Object System.Data.SqlClient.SqlConnection $cs
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = "DELETE FROM [$table] WHERE PatID = @id;
INSERT INTO [$table] (CaseID, PatID, First_Name, Middle_Name, Last_Name, PatSex, Height, Weight, Age, DOBYear, DOBMonth, DOBDay)
VALUES (@case, @id, @prenom, '', @nom, @sexe, NULL, NULL, @age, @a, @m, @j)"
        [void]$cmd.Parameters.AddWithValue('@id', $g['3000'])
        [void]$cmd.Parameters.AddWithValue('@case', ($g['3000'] + '-' + (Get-Date -Format 'yyyyMMddHHmmss')))
        [void]$cmd.Parameters.AddWithValue('@prenom', [string]$g['3102'])
        [void]$cmd.Parameters.AddWithValue('@nom', [string]$g['3101'])
        [void]$cmd.Parameters.AddWithValue('@sexe', $sexe)
        foreach ($p in @(@('@age', $age), @('@a', $a), @('@m', $m), @('@j', $j))) {
            if ($null -eq $p[1]) { [void]$cmd.Parameters.AddWithValue($p[0], [DBNull]::Value) } else { [void]$cmd.Parameters.AddWithValue($p[0], $p[1]) }
        }
        [void]$cmd.ExecuteNonQuery()
    } finally { $cn.Close() }
}

$derniereEcriture = $null
function PontGdtSql() {
    if ((LireIni $ini 'ECG' 'SqlActif' '0') -ne '1') { return }
    $dossier = LireIni $ini 'ECG' 'DossierGdt' ''
    if ($dossier.Length -eq 0) { return }
    $f = Join-Path $dossier 'IMPORT.GDT'
    if (-not (Test-Path $f)) { return }
    $t = (Get-Item $f).LastWriteTimeUtc
    if ($script:derniereEcriture -ne $null -and $t -le $script:derniereEcriture) { return }
    $script:derniereEcriture = $t
    Start-Sleep -Milliseconds 300     # fin d'ecriture par l'autre poste
    try {
        $g = LireGdt $f
        EcrireSql $g
        Journal "patinfo <- $($g['3000']) $($g['3101']) $($g['3102']) sexe=$($g['3110']) ddn=$($g['3103'])"
    } catch { Journal "ERREUR pont SQL : $($_.Exception.Message)" }
}

while ($true) {
    PontGdtSql
    # relu a chaque tour : un changement de reglage est pris en compte sans relance
    $titre   = LireIni $ini 'ECG' 'FenetreTitre' ''
    $touches = LireIni $ini 'ECG' 'TouchesValidation' '{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{UP}{DOWN}'
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
