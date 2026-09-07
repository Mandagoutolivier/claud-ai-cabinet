# build.ps1 - Construit Cabinet.dotm (Word) et Cabinet.xlsm (Excel)
# a partir des sources texte de Src\ via COM + VBIDE.
# Prerequis : "Acces approuve au modele d'objet du projet VBA" coche
# dans Word et/ou Excel (Fichier > Options > Centre de gestion de la
# confidentialite > Parametres des macros).
param(
    [ValidateSet('Word', 'Excel', 'Tous')][string]$Cible = 'Tous',
    [string]$OutDir = '',
    [switch]$Interne    # execution reelle (utilisee par le superviseur)
)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

# ---------------------------------------------------------------------
# Superviseur : l'automatisation Word/Excel peut rester bloquee de facon
# aleatoire (complements COM du poste). On execute donc le build dans un
# processus enfant surveille, tue et relance en cas de depassement.
# ---------------------------------------------------------------------
if (-not $Interne) {
    # Purge des instances Office CACHEES (residus d'automatisations passees),
    # y compris celles dont seul l'editeur VBA est devenu visible apres une
    # erreur ; les documents visibles de l'utilisateur ne sont pas touches.
    foreach ($p in @(Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue)) {
        if ($p.MainWindowHandle -eq 0 -or $p.MainWindowTitle -match '^Microsoft Visual Basic') {
            try { $p.Kill() } catch {}
        }
    }
    Start-Sleep -Milliseconds 500
    $ok = $false
    foreach ($essai in 1..3) {
        $avant = @(Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        $args = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-Cible', $Cible, '-Interne')
        if ($OutDir) { $args += @('-OutDir', "`"$OutDir`"") }
        $p = Start-Process powershell -ArgumentList $args -NoNewWindow -PassThru
        $null = $p.Handle   # indispensable pour que ExitCode soit renseigne
        if ($p.WaitForExit(180000)) {
            $p.WaitForExit()
            if ($null -eq $p.ExitCode -or $p.ExitCode -eq 0) { $ok = $true; break }
            Write-Warning "Essai $essai : echec (code $($p.ExitCode))."
            break   # vraie erreur : inutile de relancer
        }
        Write-Warning "Essai $essai : blocage au-dela de 180 s, arret et relance..."
        try { $p.Kill() } catch {}
        Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $avant -notcontains $_.Id } |
            ForEach-Object { try { $_.Kill() } catch {} }
        Start-Sleep -Seconds 2
    }
    if (-not $ok) { Write-Error 'Construction en echec apres 3 essais.' }
    exit 0
}

$DevRoot = Split-Path $PSScriptRoot -Parent
$SrcRoot = Join-Path $DevRoot 'Src'
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot 'out' }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$manifest = Get-Content (Join-Path $SrcRoot 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json

function Test-AccessVBOM([string]$app) {
    $p = "HKCU:\Software\Microsoft\Office\16.0\$app\Security"
    try { $v = (Get-ItemProperty $p -ErrorAction Stop).AccessVBOM } catch { $v = $null }
    return ($v -eq 1)
}

# Les sources sont en UTF-8 ; le VBIDE importe en ANSI (cp1252) -> transcodage
function Convert-ToAnsi([string]$path) {
    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    $tmp = Join-Path $env:TEMP ('vbasrc_' + [IO.Path]::GetFileName($path))
    [IO.File]::WriteAllText($tmp, $raw, [Text.Encoding]::GetEncoding(1252))
    return $tmp
}

function Import-Modules($vbproj, $modules) {
    foreach ($m in $modules) {
        $p = Join-Path $SrcRoot ($m -replace '/', '\')
        if (-not (Test-Path $p)) { Write-Warning "Source absente, ignoree : $m"; continue }
        $tmp = Convert-ToAnsi $p
        [void]$vbproj.VBComponents.Import($tmp)
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        Write-Host "  module  $m"
    }
}

function Set-CtlProp($ctl, [string]$prop, $value) {
    try { $ctl.$prop = $value } catch {
        try { $ctl.Object.$prop = $value } catch { Write-Warning "    propriete $prop non appliquee" }
    }
}

function Build-Form($vbproj, [string]$name) {
    $specPath = Join-Path $SrcRoot "Forms\$name.json"
    $codePath = Join-Path $SrcRoot "Forms\$name.vba"
    if (-not (Test-Path $specPath)) { Write-Warning "Spec de formulaire absente, ignoree : $name"; return }
    $spec = Get-Content $specPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $comp = $vbproj.VBComponents.Add(3)   # vbext_ct_MSForm
    $comp.Name = [string]$name
    # IMPORTANT : toujours passer aux objets COM des types .NET purs
    # ([string]/[double]) — une valeur enveloppee PSObject issue du JSON
    # provoque des dialogues invisibles et des blocages.
    $comp.Properties.Item('Caption').Value = [string]$spec.caption
    $comp.Properties.Item('Width').Value = [double]$spec.width
    $comp.Properties.Item('Height').Value = [double]$spec.height
    foreach ($c in $spec.controls) {
        $ctl = $comp.Designer.Controls.Add([string]"Forms.$($c.type).1", [string]$c.name, $true)
        $ctl.Left = [double]$c.left; $ctl.Top = [double]$c.top
        $ctl.Width = [double]$c.width; $ctl.Height = [double]$c.height
        foreach ($prop in $c.PSObject.Properties) {
            $v = $prop.Value
            switch ($prop.Name) {
                'type' {} 'name' {} 'left' {} 'top' {} 'width' {} 'height' {}
                'caption'    { Set-CtlProp $ctl 'Caption' ([string]$v) }
                'fontsize'   { try { $ctl.Font.Size = [double]$v } catch { try { $ctl.Object.Font.Size = [double]$v } catch {} } }
                'bold'       { try { $ctl.Font.Bold = [bool]$v } catch { try { $ctl.Object.Font.Bold = [bool]$v } catch {} } }
                'multiline'  { Set-CtlProp $ctl 'MultiLine' ([bool]$v) }
                'wordwrap'   { Set-CtlProp $ctl 'WordWrap' ([bool]$v) }
                'scrollbars' { Set-CtlProp $ctl 'ScrollBars' ([int]$v) }
                'style'      { Set-CtlProp $ctl 'Style' ([int]$v) }
                'textalign'  { Set-CtlProp $ctl 'TextAlign' ([int]$v) }
                'columncount'{ Set-CtlProp $ctl 'ColumnCount' ([int]$v) }
                'columnwidths' { Set-CtlProp $ctl 'ColumnWidths' ([string]$v) }
                'value'      { Set-CtlProp $ctl 'Value' $v }
                'visible'    { Set-CtlProp $ctl 'Visible' ([bool]$v) }
                'enabled'    { Set-CtlProp $ctl 'Enabled' ([bool]$v) }
                'tabindex'   { Set-CtlProp $ctl 'TabIndex' ([int]$v) }
                'default'    { Set-CtlProp $ctl 'Default' ([bool]$v) }
                'cancel'     { Set-CtlProp $ctl 'Cancel' ([bool]$v) }
                default      { Set-CtlProp $ctl $prop.Name $v }
            }
        }
    }
    if (Test-Path $codePath) {
        $code = Get-Content $codePath -Raw -Encoding UTF8
        [void]$comp.CodeModule.AddFromString($code)
    }
    Write-Host "  form    $name ($($spec.controls.Count) controles)"
}

function Build-Word {
    if (-not (Test-AccessVBOM 'Word')) {
        Write-Error "Word : l'acces au modele d'objet VBA n'est pas autorise. Cochez la case dans Fichier > Options > Centre de gestion de la confidentialite > Parametres du Centre de gestion > Parametres des macros."
    }
    Write-Host '=== Construction Cabinet.dotm (Word) ==='
    $word = New-Object -ComObject Word.Application
    Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] instance Word creee"
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        foreach ($ai in @($word.COMAddIns)) { try { $ai.Connect = $false } catch {} }
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] complements deconnectes"
        $doc = $word.Documents.Add()
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] document cree"
        # 1) Conversion (le Normal.dotm du poste peut imposer le mode de
        #    compatibilite) puis enregistrement du .dotm VIDE : un SaveAs2
        #    apres import des modules bloque Word indefiniment.
        if ($doc.CompatibilityMode -lt 15) { $doc.Convert() }
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] converti"
        [string]$out = Join-Path $OutDir ([string]$manifest.word.output)
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] out=[$($out.GetType().FullName)] '$out'"
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] cible purgee, SaveAs2..."
        $doc.SaveAs2($out, 15)    # wdFormatXMLTemplateMacroEnabled
        Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] dotm vide enregistre"
        # 2) Import des sources puis simple Save
        $vbproj = $doc.VBProject
        # Nom de projet UNIQUE : un autre modele global portant le nom par
        # defaut (TemplateProject) rendrait les cibles des raccourcis ambigues
        $vbproj.Name = 'CabinetCardio'
        Import-Modules $vbproj $manifest.word.modules
        foreach ($f in $manifest.word.forms) { Build-Form $vbproj $f }
        # 3) Raccourcis clavier enregistres DANS le modele (utilises aussi
        #    par les commandes vocales Dragon) : Ctrl+Alt+lettre.
        #    Correction = Ctrl+Alt+MAJ+C (Ctrl+Alt+C reste a l'ancien complement
        #    ModeleCourrierChatGPT_PROD.dotm, encore utilise par le medecin).
        $word.CustomizationContext = $doc
        foreach ($kb in @(
                @{ k = 78;  m = 'NouveauCourrier' },    # Ctrl+Alt+N
                @{ k = 67;  m = 'CorrigerCourrier'; maj = $true },   # Ctrl+Alt+Maj+C
                @{ k = 68;  m = 'LettreDerivee' },      # Ctrl+Alt+D
                @{ k = 80;  m = 'InsererPatient' },     # Ctrl+Alt+P
                @{ k = 117; m = 'InsererPatient'; brut = $true },   # F6 (identite du patient)
                @{ k = 71;  m = 'EnvoyerECG' },         # Ctrl+Alt+G
                @{ k = 86;  m = 'ValiderCourrier' },    # Ctrl+Alt+V
                @{ k = 66;  m = 'MettreEnGras' },       # Ctrl+Alt+B
                @{ k = 123; m = 'SondeRaccourci' })) {  # Ctrl+Alt+F12 (test fonctionnel)
            if ($kb.brut) { $code = $kb.k } else { $code = 512 + 1024 + $kb.k }
            if ($kb.maj) { $code += 256 }   # wdKeyShift
            try { [void]$word.KeyBindings.Add(2, [string]$kb.m, $code) } catch { Write-Warning "raccourci $($kb.m) non pose" }
        }
        foreach ($b in @($word.KeyBindings)) {
            if ($b.KeyString -match 'Ctrl|F6') { Write-Host "  raccourci $($b.KeyString) -> $($b.Command)" }
        }
        $doc.Save()
        $doc.Close(0)
        Write-Host "OK -> $out"
    } finally {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    # 4) Ruban "Cabinet" : injection de customUI14.xml dans le paquet OOXML
    if ($manifest.word.ruban) { Add-Ruban $out (Join-Path $SrcRoot ([string]$manifest.word.ruban)) }
}

# Ajoute (ou remplace) la partie customUI/customUI14.xml et sa relation dans
# un .dotm deja ferme. Word charge le ruban des modeles globaux du STARTUP.
function Add-Ruban([string]$dotm, [string]$xml) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($dotm, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $e = $zip.GetEntry('customUI/customUI14.xml'); if ($e) { $e.Delete() }
        $e = $zip.CreateEntry('customUI/customUI14.xml')
        $s = $e.Open(); $b = [IO.File]::ReadAllBytes($xml); $s.Write($b, 0, $b.Length); $s.Dispose()
        $rels = $zip.GetEntry('_rels/.rels')
        $r = New-Object IO.StreamReader($rels.Open()); $txt = $r.ReadToEnd(); $r.Dispose()
        if ($txt -notmatch 'customUI14\.xml') {
            $txt = $txt -replace '</Relationships>', '<Relationship Id="rIdCabinetUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>'
            $rels.Delete()
            $rels = $zip.CreateEntry('_rels/.rels')
            $w = New-Object IO.StreamWriter($rels.Open(), (New-Object Text.UTF8Encoding $false)); $w.Write($txt); $w.Dispose()
        }
        $ct = $zip.GetEntry('[Content_Types].xml')
        $r = New-Object IO.StreamReader($ct.Open()); $ctx = $r.ReadToEnd(); $r.Dispose()
        if ($ctx -notmatch 'Extension="xml"') {
            $ctx = $ctx -replace '<Default Extension="rels"', '<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels"'
            $ct.Delete(); $ct = $zip.CreateEntry('[Content_Types].xml')
            $w = New-Object IO.StreamWriter($ct.Open(), (New-Object Text.UTF8Encoding $false)); $w.Write($ctx); $w.Dispose()
        }
        Write-Host "  ruban Cabinet injecte ($([IO.Path]::GetFileName($xml)))"
    } finally { $zip.Dispose() }
}

function Build-Excel {
    if (-not (Test-AccessVBOM 'Excel')) {
        Write-Error "Excel : l'acces au modele d'objet VBA n'est pas autorise. Ouvrez Excel : Fichier > Options > Centre de gestion de la confidentialite > Parametres du Centre de gestion > Parametres des macros > cochez 'Acces approuve au modele d'objet du projet VBA', puis relancez ce script."
    }
    Write-Host '=== Construction Cabinet.xlsm (Excel) ==='
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        foreach ($ai in @($excel.COMAddIns)) { try { $ai.Connect = $false } catch {} }
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        # Enregistrement du classeur VIDE d'abord (meme precaution que Word)
        [string]$out = Join-Path $OutDir ([string]$manifest.excel.output)
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
        $wb.SaveAs($out, 52)      # xlOpenXMLWorkbookMacroEnabled
        $vbproj = $wb.VBProject
        Import-Modules $vbproj $manifest.excel.modules
        foreach ($f in $manifest.excel.forms) { Build-Form $vbproj $f }
        if ($manifest.excel.thisworkbook) {
            $twPath = Join-Path $SrcRoot ($manifest.excel.thisworkbook -replace '/', '\')
            if (Test-Path $twPath) {
                $code = Get-Content $twPath -Raw -Encoding UTF8
                [void]$vbproj.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString($code)
                Write-Host '  code    ThisWorkbook'
            }
        }
        # Feuille Accueil avec boutons
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Accueil'
        $ws.Cells.Item(1, 2).Value2 = 'Gestion du cabinet'
        $ws.Cells.Item(1, 2).Font.Size = 18
        $ws.Cells.Item(1, 2).Font.Bold = $true
        $ws.Cells.Item(2, 2).Value2 = 'Cliquez sur une action ci-dessous.'
        # ligne 4 = bandeau "courriers en attente" (ecrit par modEchange) :
        # ligne haute, colonne large, boutons places nettement en dessous
        $ws.Cells.Item(4, 2).Value2 = 'Aucun courrier en attente.'
        $ws.Rows.Item(4).RowHeight = 30
        $ws.Cells.Item(4, 2).VerticalAlignment = -4108   # centre
        $ws.Columns.Item(2).ColumnWidth = 70
        $top = 110.0
        foreach ($b in $manifest.excel.boutons) {
            $btn = $ws.Buttons().Add(60.0, $top, 220.0, 30.0)
            $btn.Text = [string]$b.texte
            $btn.OnAction = [string]$b.macro
            $top += 40.0
        }
        $ws.Columns.Item(1).ColumnWidth = 3
        # Feuilles supplementaires (ex : Agenda) avec boutons et code evenementiel
        foreach ($f in @($manifest.excel.feuilles)) {
            $wsF = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
            $wsF.Name = [string]$f.nom
            # boutons par rangees de 5 (au-dela, une 2e rangee dans la meme
            # ligne 1, dont la hauteur est doublee : la grille commence ligne 2)
            $parRangee = 5
            $nb = @($f.boutons).Count
            $rangees = [math]::Ceiling($nb / $parRangee)
            $wsF.Rows.Item(1).RowHeight = 4.0 + 28.0 * [math]::Max(1, $rangees)
            $i = 0
            foreach ($b in @($f.boutons)) {
                $left = 10.0 + 158.0 * ($i % $parRangee)
                $top = 4.0 + 28.0 * [math]::Floor($i / $parRangee)
                $btn = $wsF.Buttons().Add($left, $top, 150.0, 24.0)
                $btn.Text = [string]$b.texte
                $btn.OnAction = [string]$b.macro
                $i++
            }
            if ($f.code) {
                $codePath = Join-Path $SrcRoot (([string]$f.code) -replace '/', '\')
                if (Test-Path $codePath) {
                    $wb.Save()   # attribue le nom de code (CodeName) a la nouvelle feuille
                    $compF = $null
                    $cn = [string]$wsF.CodeName
                    if ($cn) { try { $compF = $vbproj.VBComponents.Item($cn) } catch {} }
                    if (-not $compF) {
                        foreach ($c in @($vbproj.VBComponents)) {
                            $nomProp = ''
                            try { $nomProp = [string]$c.Properties.Item('Name').Value } catch {}
                            Write-Host "    composant $($c.Name) type=$([int]$c.Type) nom=$nomProp"
                            if ($nomProp -eq [string]$f.nom) { $compF = $c }
                        }
                    }
                    if ($compF) { [void]$compF.CodeModule.AddFromString((Get-Content $codePath -Raw -Encoding UTF8)); Write-Host "  code    feuille $($f.nom) ($($compF.Name))" }
                    else { Write-Warning "module de feuille introuvable pour $($f.nom)" }
                }
            }
            Write-Host "  feuille $($f.nom) ($(@($f.boutons).Count) boutons)"
        }
        $wb.Worksheets.Item('Accueil').Activate()
        $wb.Save()
        $wb.Close($false)
        Write-Host "OK -> $out"
    } finally {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

if ($Cible -eq 'Word' -or $Cible -eq 'Tous') { Build-Word }
if ($Cible -eq 'Excel' -or $Cible -eq 'Tous') { Build-Excel }
Write-Host 'Construction terminee.'
