# init_donnees.ps1 - Cree l'arborescence de donnees du cabinet et les
# classeurs de base (Patients, Agenda, Journal, Nomenclature).
# -Racine : dossier de donnees (defaut : SandboxData du depot de dev)
# -Sandbox : injecte des patients FICTIFS pour les tests
# -Force : recree les classeurs meme s'ils existent (ecrase !)
param(
    [string]$Racine = '',
    [switch]$Sandbox,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

$DevRoot = Split-Path $PSScriptRoot -Parent
if (-not $Racine) { $Racine = Join-Path $DevRoot 'SandboxData' }
$annee = (Get-Date).Year

Write-Host "Initialisation des donnees dans : $Racine"
foreach ($d in 'Base\locks', 'Actes', 'Patients', 'Echange\AEnvoyer', 'Echange\Traites',
               'Modeles\Deploy', 'Config\prompts', 'Config\style', 'Sauvegardes', 'Logs') {
    New-Item -ItemType Directory -Force -Path (Join-Path $Racine $d) | Out-Null
}

# --- Config par defaut (copie sans ecraser, sauf -Force) --------------
$cfgSrc = Join-Path $DevRoot 'Src\ConfigDefaut'
if (Test-Path $cfgSrc) {
    Get-ChildItem $cfgSrc -Recurse -File | ForEach-Object {
        $rel = $_.FullName.Substring($cfgSrc.Length + 1)
        $dest = Join-Path (Join-Path $Racine 'Config') $rel
        New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
        if ($Force -or -not (Test-Path $dest)) {
            Copy-Item $_.FullName $dest -Force
            Write-Host "  config : $rel"
        }
    }
}

# --- Classeurs --------------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function New-Classeur([string]$path, $feuilles) {
    if ((Test-Path $path) -and -not $Force) { Write-Host "  existe deja : $path"; return }
    if (Test-Path $path) { Remove-Item $path -Force }
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    for ($i = 0; $i -lt $feuilles.Count; $i++) {
        if ($i -eq 0) { $ws = $wb.Worksheets.Item(1) }
        else { $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
        $f = $feuilles[$i]
        $ws.Name = $f.Nom
        if ($f.ColonnesTexte) {
            foreach ($ct in $f.ColonnesTexte) { $ws.Columns.Item($ct).NumberFormat = '@' }
        }
        $col = 1
        foreach ($h in $f.Entetes) { $ws.Cells.Item(1, $col).Value2 = $h; $col++ }
        $ws.Rows.Item(1).Font.Bold = $true
        $r = 2
        if ($f.Lignes) {
            foreach ($ligne in $f.Lignes) {
                $col = 1
                foreach ($v in $ligne) { $ws.Cells.Item($r, $col).Value2 = $v; $col++ }
                $r++
            }
        }
        [void]$ws.Columns.AutoFit()
    }
    $wb.SaveAs($path, 51)   # xlsx
    $wb.Close($false)
    Write-Host "  cree : $path"
}

try {
    # PATIENTS + CORRESPONDANTS
    $entPat = @('ID','Nom','NomNaissance','Prenom','DDN','Sexe','NIR','Adresse1','Adresse2','CP','Ville','Tel','Mobile','Email','MedTraitantID','Mutuelle','ALD','Notes','DateCreation','DateModif')
    $entCor = @('ID','Titre','Nom','Prenom','Specialite','Adresse1','Adresse2','CP','Ville','Tel','Email','FormuleAppel','FormulePolitesse','Actif')
    $lgPat = @(); $lgCor = @()
    if ($Sandbox) {
        $lgCor = @(
            ,@('C0001','Dr','GARRIGUE','Paul','Medecine generale','12 avenue des Platanes','','34000','Montpellier','04 67 00 00 01','','Cher Ami,','Bien confraternellement.','1')
            ,@('C0002','Dr','VALLESPIR','Marie','Medecine generale','3 rue du Marche','','34170','Castelnau-le-Lez','04 67 00 00 02','','Chere Amie,','Bien confraternellement.','1')
            ,@('C0003','Pr','DELMAS','Antoine','Cardiologie interventionnelle','CHU - Service de cardiologie','371 avenue du Doyen Giraud','34295','Montpellier','04 67 00 00 03','','Cher Ami,','Bien amicalement.','1')
            ,@('C0004','Dr','BONNAFOUS','Claire','Pneumologie','8 place de la Comedie','','34000','Montpellier','04 67 00 00 04','','Chere Consoeur,','Confraternellement.','1')
            ,@('C0005','Dr','RIVIERE','Luc','Nephrologie','25 boulevard Pasteur','','34500','Beziers','04 67 00 00 05','','Cher Confrere,','Confraternellement.','1')
        )
        $prenoms = @('Jean','Marie','Pierre','Jeanne','Michel','Francoise','Andre','Monique','Rene','Denise','Paul','Yvette','Louis','Simone','Georges','Odette','Marcel','Therese','Roger','Colette')
        $noms    = @('FABREGUE','SANTONJA','COMBALUZIER','PLAGNOL','ROUVIERE','ESCANDE','BASTIDE','CAUSSE','MAURIN','SEGURA','ALIBERT','VIDALOU','TRINQUIER','BOSC','GINESTET','LAFON','SALVAGNAC','PORTAL','MAZEL','CANTAGREL')
        for ($i = 0; $i -lt 20; $i++) {
            # ATTENTION : ne jamais utiliser d'operateur (+, -f) DANS un litteral
            # de tableau PowerShell (la virgule est prioritaire) - tout precalculer.
            $id = 'P' + ('{0:d5}' -f ($i + 1))
            $sexe = @('M','F')[$i % 2]
            $jour = ('{0:d2}' -f (1 + ($i * 3) % 28)); $mois = ('{0:d2}' -f (1 + $i % 12)); $an = 1935 + ($i * 2)
            $ddn = "$jour/$mois/$an"
            $nir = ('{0}{1}{2}34000{3:d4}' -f @('1','2')[$i % 2], ($an % 100).ToString('d2'), $mois, ($i + 1)) + ' 42'
            $med = 'C000' + (1 + $i % 2)
            $i2 = '{0:d2}' -f $i
            $tel = "04 67 11 22 $i2"
            $mob = "06 12 34 56 $i2"
            $adr = "$($i + 1) rue des Oliviers"
            $ald = @('','O')[$i % 2]
            $auj = Get-Date -Format 'dd/MM/yyyy'
            $lgPat += ,@($id, $noms[$i], '', $prenoms[$i], $ddn, $sexe, $nir, $adr, '', '34000', 'Montpellier', $tel, $mob, '', $med, '', $ald, '', $auj, '')
        }
    }
    New-Classeur (Join-Path $Racine 'Base\Patients.xlsx') @(
        @{ Nom = 'PATIENTS'; Entetes = $entPat; Lignes = $lgPat; ColonnesTexte = @('A','E','G','J','L','M','Q','S','T') },
        @{ Nom = 'CORRESPONDANTS'; Entetes = $entCor; Lignes = $lgCor; ColonnesTexte = @('A','J','N') }
    )

    # AGENDA
    $entRdv = @('ID','PatientID','Date','Heure','DureeMin','TypeActe','Statut','HeureArrivee','Notes','DateCreation')
    $lgRdv = @()
    if ($Sandbox) {
        $auj = Get-Date -Format 'dd/MM/yyyy'
        $lgRdv = @(
            ,@('R00001','P00001',$auj,'09:00','30','CSC','Prevu','','','')
            ,@('R00002','P00002',$auj,'09:30','30','CSC','Prevu','','','')
            ,@('R00003','P00003',$auj,'10:00','45','ETT','Prevu','','','')
        )
    }
    New-Classeur (Join-Path $Racine "Base\Agenda_$annee.xlsx") @(
        @{ Nom = 'RDV'; Entetes = $entRdv; Lignes = $lgRdv; ColonnesTexte = @('A','B','C','D','H','J') }
    )

    # JOURNAL COMPTABLE
    $entJnl = @('Date','SeanceID','PatientID','Nom','Prenom','DDN','NIR','CodeActe','Montant','ModePaiement','TiersPayant','Paye','DateEncaissement','FeuilleSoinsImprimee','Notes')
    New-Classeur (Join-Path $Racine "Actes\Journal_$annee.xlsx") @(
        @{ Nom = 'JOURNAL'; Entetes = $entJnl; ColonnesTexte = @('A','B','C','F','G','M') }
    )

    # NOMENCLATURE (valeurs renseignees par le cabinet le 02/09/2026 - a tenir a jour)
    $entAct = @('Code','LibelleCourt','LibelleCerfa','Tarif','CodeAssocie','TarifAssocie','Depassement','Actif')
    $lgAct = @(
        ,@('CS',      'Consultation simple',                  'CS',      '26.5',  'MCS',     '5',     '0', '1')
        ,@('CSC',     'Consultation specifique cardio + ECG', 'CSC',     '47.73', 'MCC',     '4.77',  '0', '1')
        ,@('ETT',     'Echocardiographie (ETT)',              'DZQM006', '94.28', '',        '0',     '0', '1')
        ,@('HOLTER',  'Holter ECG',                           'DEQP005', '77.01', '',        '0',     '0', '1')
        ,@('MAPA',    'MAPA',                                 'EQQP008', '50',    '',        '0',     '0', '1')
        ,@('APC',     'Avis ponctuel de consultant',          'APC',     '60',    'DEQP003', '14.52', '0', '1')
    )
    New-Classeur (Join-Path $Racine 'Config\Nomenclature.xlsx') @(
        @{ Nom = 'ACTES'; Entetes = $entAct; Lignes = $lgAct; ColonnesTexte = @('A','B','C','E') }
    )

    # DICTIONNAIRES DE MISE EN GRAS (graines : Src\ConfigDefaut\gras\*.txt)
    function Lire-Liste([string]$f) {
        if (-not (Test-Path $f)) { return @() }
        return @(Get-Content $f -Encoding UTF8 | Where-Object { $_.Trim() -ne '' -and -not $_.Trim().StartsWith('#') })
    }
    $lgMed = @()
    foreach ($l in (Lire-Liste (Join-Path $cfgSrc 'gras\medicaments.txt'))) {
        $p = $l -split '\|', 2
        $var = ''; if ($p.Count -gt 1) { $var = $p[1].Trim() }
        $lgMed += ,@($p[0].Trim(), $var, '1')
    }
    $lgExp = @()
    foreach ($l in (Lire-Liste (Join-Path $cfgSrc 'gras\expressions.txt'))) { $lgExp += ,@($l.Trim(), '1') }
    New-Classeur (Join-Path $Racine 'Config\Gras_Medicaments.xlsx') @(
        @{ Nom = 'MEDICAMENTS'; Entetes = @('Terme','Variantes','Actif'); Lignes = $lgMed; ColonnesTexte = @('A','B','C') }
    )
    New-Classeur (Join-Path $Racine 'Config\Gras_Expressions.xlsx') @(
        @{ Nom = 'EXPRESSIONS'; Entetes = @('Terme','Actif'); Lignes = $lgExp; ColonnesTexte = @('A','B') }
    )
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
Write-Host 'Initialisation terminee.'
if ($Sandbox) { Write-Host 'ATTENTION : donnees FICTIVES (sandbox) - ne pas utiliser en production.' }
