# installer_relais.ps1 - Installe le module-relais et les raccourcis
# Ctrl+Alt+N/D/P/G/V/B et Ctrl+Alt+Maj+C dans Normal.dotm du poste
# (voir modCabinetRelais.bas). Ctrl+Alt+C est LIBERE : il appartient a
# l'ancien complement ModeleCourrierChatGPT_PROD.dotm, encore utilise.
# Prerequis : Word FERME (Normal.dotm doit etre modifiable) et acces VBA
# autorise ("Acces approuve au modele d'objet du projet VBA").
# Idempotent : reinstalle le module et les raccourcis a chaque appel.
param([switch]$Test)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')

$src = Join-Path (Split-Path $PSScriptRoot -Parent) 'Src\Word\Relais\modCabinetRelais.bas'
if (-not (Test-Path $src)) { Write-Error "Source introuvable : $src" }
if (Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }) {
    Write-Error 'Word est ouvert : fermez-le completement avant d''installer le relais.'
}
Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch {} }
Start-Sleep -Seconds 1

# transcodage UTF-8 -> ANSI pour l'import VBIDE
$tmp = Join-Path $env:TEMP 'modCabinetRelais.bas'
[IO.File]::WriteAllText($tmp, (Get-Content $src -Raw -Encoding UTF8), [Text.Encoding]::GetEncoding(1252))

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $word.AutomationSecurity = 1
    Start-Sleep -Seconds 2
    [void]$word.Documents.Add()
    $normal = $word.NormalTemplate
    $proj = $normal.VBProject
    # remplacement du module existant
    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCabinetRelais') { $proj.VBComponents.Remove($c) } }
    [void]$proj.VBComponents.Import($tmp)
    Write-Host "module modCabinetRelais installe dans $($normal.FullName)"

    $word.CustomizationContext = $normal
    $liaisons = @(
        @{ k = 78;  m = 'CAB_NouveauCourrier' },   # Ctrl+Alt+N
        @{ k = 67;  m = 'CAB_CorrigerCourrier'; maj = $true },  # Ctrl+Alt+Maj+C
        @{ k = 68;  m = 'CAB_LettreDerivee' },     # Ctrl+Alt+D
        @{ k = 80;  m = 'CAB_InsererPatient' },    # Ctrl+Alt+P
        @{ k = 117; m = 'CAB_InsererPatient'; brut = $true },   # F6
        @{ k = 71;  m = 'CAB_EnvoyerECG' },        # Ctrl+Alt+G
        @{ k = 86;  m = 'CAB_ValiderCourrier' },   # Ctrl+Alt+V
        @{ k = 66;  m = 'CAB_MettreEnGras' },      # Ctrl+Alt+B
        @{ k = 65;  m = 'CAB_AllerDestinataire'; maj = $true },  # Ctrl+Alt+Maj+A
        @{ k = 66;  m = 'CAB_AllerAppel'; maj = $true },         # Ctrl+Alt+Maj+B
        @{ k = 68;  m = 'CAB_FinaliserCourrier'; maj = $true },  # Ctrl+Alt+Maj+D
        @{ k = 123; m = 'CAB_Sonde' })             # Ctrl+Alt+F12 (test)
    # liberer Ctrl+Alt+C si une version precedente du relais l'avait pris
    try {
        $ancien = $word.FindKey(512 + 1024 + 67)
        if ($ancien.Command -match 'CAB_CorrigerCourrier') { $ancien.Clear(); Write-Host '  Ctrl+Alt+C libere (rendu a l''ancien complement)' }
    } catch {}
    foreach ($l in $liaisons) {
        if ($l.brut) { $code = $l.k } else { $code = 512 + 1024 + $l.k }
        if ($l.maj) { $code += 256 }   # wdKeyShift
        try { $word.FindKey($code).Clear() } catch {}
        [void]$word.KeyBindings.Add(2, [string]$l.m, $code)
        Write-Host ("  {0,-14} -> {1}" -f $word.FindKey($code).KeyString, $word.FindKey($code).Command)
    }
    $normal.Save()
    Write-Host 'Normal.dotm enregistre.'

    # test fonctionnel : la liaison Ctrl+Alt+F12 doit ecrire le fichier temoin
    $temoin = Join-Path $env:TEMP 'CabinetCardio\sonde_raccourci.txt'
    Remove-Item $temoin -Force -ErrorAction SilentlyContinue
    $word.FindKey(512 + 1024 + 123).Execute()
    Start-Sleep -Milliseconds 800
    if (Test-Path $temoin) { Write-Host 'TEST OK : le raccourci relais declenche bien la macro de Cabinet.dotm.' }
    else { Write-Warning 'TEST ECHEC : le raccourci relais n''a pas declenche la macro (Cabinet.dotm charge ?).' }
    Remove-Item $temoin -Force -ErrorAction SilentlyContinue
} finally {
    try { $word.Quit(0) } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}
