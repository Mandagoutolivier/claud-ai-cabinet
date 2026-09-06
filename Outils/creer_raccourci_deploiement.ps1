# =====================================================================
# creer_raccourci_deploiement.ps1 - A executer UNE FOIS, depuis un poste
# qui voit le NAS (ex : ce PC du domicile). Depose sur le NAS, dans un
# dossier dedie, un raccourci "Deployer le cabinet.lnk" pret a etre copie
# sur le Bureau du poste medecin au cabinet.
#
# Le raccourci ne fait AUCUNE hypothese de version : au double-clic, il
# telecharge la DERNIERE version de deployer_cabinet.ps1 depuis GitHub
# (ou, a defaut d'internet, reprend la copie du NAS) puis l'execute.
# Vous n'aurez donc plus jamais a regenerer ce raccourci pour profiter
# des futures corrections.
# =====================================================================
param(
    [string]$Nas     = '\\DS224\home\claude\claude ai',
    [string]$Branche = 'claude/suivi-dev-logiciel-cabinet-fdjpa9'
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $Nas)) { throw "NAS inaccessible : $Nas" }

$dossier = Join-Path $Nas 'Deploiement-Cabinet'
New-Item -ItemType Directory -Force -Path $dossier | Out-Null

$urlBrute = "https://raw.githubusercontent.com/Mandagoutolivier/claud-ai-cabinet/$Branche/Outils/deployer_cabinet.ps1"
$copieNas = Join-Path $Nas 'CabinetCardio-Git\Outils\deployer_cabinet.ps1'

# Cette commande est volontairement COURTE et STABLE : c'est elle qui est
# figee dans le raccourci. Toute amelioration future de deployer_cabinet.ps1
# sur GitHub sera prise en compte automatiquement, sans regenerer le lnk.
$commande = @"
`$ErrorActionPreference='Stop'; `$d=Join-Path `$env:TEMP 'deployer_cabinet.ps1';
try { Invoke-WebRequest -UseBasicParsing '$urlBrute' -OutFile `$d }
catch {
  Write-Host 'Internet indisponible : reprise de la copie du NAS (verifiez sa fraicheur).' -ForegroundColor Yellow
  Copy-Item '$copieNas' `$d -Force
}
& `$d
"@
# une seule ligne, guillemets doubles internes proteges pour l'argument -Command
$commandeUneLigne = ($commande -replace "`r?`n", ' ')

$lnk = Join-Path $dossier 'Deployer le cabinet.lnk'
$wsh = New-Object -ComObject WScript.Shell
$raccourci = $wsh.CreateShortcut($lnk)
$raccourci.TargetPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$raccourci.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"$commandeUneLigne`""
$raccourci.WorkingDirectory = $env:TEMP
$raccourci.Description = 'Deploie la derniere version du logiciel de cabinet (medecin + secretariat)'
$raccourci.IconLocation = "$env:WINDIR\System32\shell32.dll,21"
$raccourci.Save()

@"
Ce dossier contient le raccourci de deploiement du logiciel de cabinet.

A FAIRE (une fois, au cabinet, sur le PC du medecin) :
  1. Copier "Deployer le cabinet.lnk" sur le Bureau de ce PC.
  2. Fermer Word et Excel.
  3. Double-cliquer le raccourci.

Le raccourci recupere toujours la DERNIERE version depuis GitHub : vous
n'avez pas besoin de le regenerer apres une correction du logiciel.

Il installe automatiquement ce poste (medecin), et tente de mettre a jour
le poste secretariat (RDC) a distance. Si cette derniere etape echoue
(reseau, pare-feu), le script depose sur \\RDC\CabinetCardio\_Installation
un lanceur "installer_secretariat.cmd" : il suffit alors de le double-
cliquer UNE FOIS depuis le PC RDC lui-meme (Word/Excel fermes la-bas aussi).

En cas de probleme, relancez simplement le raccourci : aucune etape n'est
destructive, la base des patients n'est jamais touchee.
"@ | Out-File (Join-Path $dossier 'LISEZMOI.txt') -Encoding UTF8

Write-Host "Raccourci cree : $lnk" -ForegroundColor Green
Write-Host "Copiez-le sur le Bureau du poste medecin au cabinet, puis double-cliquez." -ForegroundColor Green
