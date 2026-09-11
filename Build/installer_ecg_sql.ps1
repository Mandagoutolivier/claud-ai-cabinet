# =====================================================================
# installer_ecg_sql.ps1 - Prepare la "boite aux lettres" SQL lue par
# Resting12Lead (connexion Systeme Info DMS) sur le poste MEDECIN.
# A lancer UNE FOIS, en administrateur, sur AX8_Max, apres installation de
# SQL Server Express (instance SQLEXPRESS) :
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\installer_ecg_sql.ps1
#
# Ce script :
#   1. verifie l'instance SQL ; si absente, tente l'installation par winget
#      (Microsoft.SQLServer.2022.Express) sinon indique le telechargement ;
#   2. active l'authentification mixte (login SQL) et redemarre l'instance ;
#   3. cree la base, le login (SetSQLInfo.ini de Resting12Lead : DMSNIS /
#      dmsdms, base ecgcenter), les tables patinfo (entree) et returninfo
#      (sortie) avec les colonnes attendues ;
#   4. ecrit une ligne de test puis la supprime.
# Les valeurs viennent de [ECG] Sql* du config.ini du poste (chemin.txt).
# Le remplissage de patinfo a chaque "Arrive" est fait par le veilleur
# ecg_valider_fenetre.ps1 ([ECG] SqlActif=1).
# =====================================================================
param(
    [string]$Instance = '',
    [string]$Base = '',
    [string]$Utilisateur = '',
    [string]$MotDePasse = ''
)
$ErrorActionPreference = 'Stop'
trap { Write-Host ''; Write-Host "ARRET : $($_.Exception.Message)" -ForegroundColor Red; Read-Host 'Appuyez sur Entree pour fermer cette fenetre'; exit 1 }
function Etape([string]$t) { Write-Host ''; Write-Host "=== $t" -ForegroundColor Cyan }
function Ok([string]$t)    { Write-Host "  OK  $t" -ForegroundColor Green }
function Info([string]$t)  { Write-Host "  --  $t" -ForegroundColor Yellow }
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
$ini = ''
if (Test-Path $cheminTxt) { $ini = Join-Path ((Get-Content $cheminTxt -TotalCount 1).Trim()) 'Config\config.ini' }
if ($Instance -eq '')    { $Instance    = LireIni $ini 'ECG' 'SqlInstance' '.\SQLEXPRESS' }
if ($Base -eq '')        { $Base        = LireIni $ini 'ECG' 'SqlBase' 'ecgcenter' }
if ($Utilisateur -eq '') { $Utilisateur = LireIni $ini 'ECG' 'SqlUtilisateur' 'DMSNIS' }
if ($MotDePasse -eq '')  { $MotDePasse  = LireIni $ini 'ECG' 'SqlMotDePasse' 'dmsdms' }
$nomInstance = if ($Instance -match '\\(.+)$') { $Matches[1] } else { 'MSSQLSERVER' }
$service = if ($nomInstance -eq 'MSSQLSERVER') { 'MSSQLSERVER' } else { "MSSQL`$$nomInstance" }

Etape "Instance SQL Server ($Instance)"
$svc = Get-Service -Name $service -ErrorAction SilentlyContinue
if (-not $svc) {
    Info "service $service introuvable : SQL Server Express n'est pas installe."
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Info 'tentative d installation par winget (quelques minutes, telechargement ~300 Mo)...'
        & winget install --id Microsoft.SQLServer.2022.Express -e --accept-package-agreements --accept-source-agreements --silent
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
    }
    if (-not $svc) {
        throw "Installez SQL Server Express (https://www.microsoft.com/fr-fr/sql-server/sql-server-downloads, edition Express, installation de base, instance SQLEXPRESS) puis relancez ce script."
    }
}
if ($svc.Status -ne 'Running') { Start-Service $service; Start-Sleep 3 }
Ok "service $service en cours d execution"

Etape 'Authentification mixte (login SQL pour Resting12Lead)'
$regBase = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server'
$idInst = (Get-ItemProperty "$regBase\Instance Names\SQL" -ErrorAction SilentlyContinue).$nomInstance
if ($idInst) {
    $regInst = "$regBase\$idInst\MSSQLServer"
    if ((Get-ItemProperty $regInst).LoginMode -ne 2) {
        Set-ItemProperty $regInst -Name LoginMode -Value 2
        Restart-Service $service -Force; Start-Sleep 5
        Ok 'mode mixte active (instance redemarree)'
    } else { Ok 'mode mixte deja actif' }
} else { Info "cle de registre de l instance introuvable : verifiez le mode mixte a la main (SSMS > Proprietes > Securite)" }

function Sql([string]$cs, [string]$texte) {
    $cn = New-Object System.Data.SqlClient.SqlConnection $cs
    $cn.Open()
    try { $c = $cn.CreateCommand(); $c.CommandText = $texte; [void]$c.ExecuteNonQuery() } finally { $cn.Close() }
}
$csAdmin = "Server=$Instance;Database=master;Integrated Security=True;Connect Timeout=10"

Etape "Base $Base, login $Utilisateur, tables patinfo / returninfo"
Sql $csAdmin "IF DB_ID('$Base') IS NULL CREATE DATABASE [$Base];"
Sql $csAdmin "IF NOT EXISTS (SELECT 1 FROM sys.sql_logins WHERE name = '$Utilisateur') CREATE LOGIN [$Utilisateur] WITH PASSWORD = '$MotDePasse', CHECK_POLICY = OFF; ELSE ALTER LOGIN [$Utilisateur] WITH PASSWORD = '$MotDePasse', CHECK_POLICY = OFF;"
$csBase = "Server=$Instance;Database=$Base;Integrated Security=True;Connect Timeout=10"
Sql $csBase "IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '$Utilisateur') CREATE USER [$Utilisateur] FOR LOGIN [$Utilisateur]; ALTER ROLE db_owner ADD MEMBER [$Utilisateur];"
Sql $csBase @"
IF OBJECT_ID('patinfo') IS NULL CREATE TABLE patinfo (
    CaseID      NVARCHAR(50)  NOT NULL PRIMARY KEY,
    PatID       NVARCHAR(50)  NOT NULL,
    First_Name  NVARCHAR(100) NULL,
    Middle_Name NVARCHAR(100) NULL,
    Last_Name   NVARCHAR(100) NULL,
    PatSex      NVARCHAR(10)  NULL,
    Height      FLOAT NULL,
    Weight      FLOAT NULL,
    Age         INT NULL,
    DOBYear     INT NULL,
    DOBMonth    INT NULL,
    DOBDay      INT NULL,
    Insere      DATETIME NOT NULL DEFAULT GETDATE()
);
IF OBJECT_ID('returninfo') IS NULL CREATE TABLE returninfo (
    DrvID            INT IDENTITY(1,1) PRIMARY KEY,
    CASE_ID          NVARCHAR(50)  NULL,
    Conclusion       NVARCHAR(MAX) NULL,
    Conclusion_Sheet NVARCHAR(MAX) NULL,
    PDF_Fname        NVARCHAR(260) NULL,
    PDF_DIR          NVARCHAR(260) NULL,
    Recu             DATETIME NOT NULL DEFAULT GETDATE()
);
"@
Ok 'base, login et tables en place'

Etape 'Test avec le login de Resting12Lead'
$csTest = "Server=$Instance;Database=$Base;User ID=$Utilisateur;Password=$MotDePasse;Connect Timeout=10"
Sql $csTest "INSERT INTO patinfo (CaseID, PatID, First_Name, Last_Name, PatSex, DOBYear, DOBMonth, DOBDay) VALUES ('TEST-INSTALL', 'TEST', 'Test', 'INSTALLATION', 'F', 1950, 1, 1); DELETE FROM patinfo WHERE PatID = 'TEST';"
Ok "ecriture / suppression d une ligne de test reussies ($Utilisateur)"

Etape 'Bilan'
Write-Host "  Dans Resting12Lead : Parametres > Parametres Interface > cocher 'Connection Systeme Info DMS'" -ForegroundColor Green
Write-Host "  et verifier SetSQLInfo.ini : DataSource=$Base (ou la source ODBC pointant sur $Instance), DBUserID=$Utilisateur." -ForegroundColor Green
Write-Host "  Dans Config\config.ini du cabinet : [ECG] SqlActif=1 (le veilleur remplit patinfo a chaque 'Arrive')." -ForegroundColor Green
Read-Host 'Appuyez sur Entree pour fermer cette fenetre'
