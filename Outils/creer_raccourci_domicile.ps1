# =====================================================================
# creer_raccourci_domicile.ps1 - A executer UNE FOIS sur le PC du domicile.
# Depose sur le Bureau un raccourci "Mettre a jour le cabinet (domicile)"
# qui, a chaque double-clic, telecharge la DERNIERE version de
# maj_poste.ps1 depuis GitHub (copie du NAS en secours) et l'execute avec
# -Role Tous : medecin + secretariat sur ce PC.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\creer_raccourci_domicile.ps1
# =====================================================================
param(
    [string]$Nas     = '\\DS224\home\claude\claude ai',
    [string]$Branche = 'claude/suivi-dev-logiciel-cabinet-fdjpa9',
    [ValidateSet('Tous', 'Medecin', 'Secretaire')][string]$Role = 'Tous'
)
$ErrorActionPreference = 'Stop'

$urlBrute = "https://raw.githubusercontent.com/Mandagoutolivier/claud-ai-cabinet/$Branche/Outils/maj_poste.ps1"
$copieNas = Join-Path $Nas 'CabinetCardio-Git\Outils\maj_poste.ps1'

# Commande courte et stable figee dans le raccourci : les ameliorations
# futures de maj_poste.ps1 sont prises en compte sans le regenerer.
$commande = @"
`$ErrorActionPreference='Stop'; `$d=Join-Path `$env:TEMP 'maj_poste.ps1';
try { Invoke-WebRequest -UseBasicParsing ('$urlBrute' + '?nocache=' + (Get-Date -Format yyyyMMddHHmmss)) -OutFile `$d }
catch {
  Write-Host 'Internet indisponible : reprise de la copie du NAS (verifiez sa fraicheur).' -ForegroundColor Yellow
  Copy-Item '$copieNas' `$d -Force
}
& `$d -Role $Role
"@
$commandeUneLigne = ($commande -replace "`r?`n", ' ')

$bureau = [Environment]::GetFolderPath('Desktop')
$lnk = Join-Path $bureau "Mettre a jour le cabinet (domicile).lnk"
$wsh = New-Object -ComObject WScript.Shell
$raccourci = $wsh.CreateShortcut($lnk)
$raccourci.TargetPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$raccourci.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"$commandeUneLigne`""
$raccourci.WorkingDirectory = $env:TEMP
$raccourci.Description = "Recupere la derniere version GitHub et installe le logiciel sur ce PC (role $Role)"
$raccourci.IconLocation = "$env:WINDIR\System32\shell32.dll,21"
$raccourci.Save()

Write-Host "Raccourci cree sur le Bureau : $lnk" -ForegroundColor Green
Write-Host "Fermez Word et Excel, puis double-cliquez-le pour mettre a jour (role $Role)." -ForegroundColor Green
