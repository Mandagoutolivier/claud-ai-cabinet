# Cabinet Cardio — logiciel de gestion (Dr Mandagout)

## 📌 Pense-bête en cours

**Déploiement au cabinet prévu lundi.** Récapitulatif (à jour au 06/09/2026) :

1. Raccourci déjà déposé sur le NAS :
   `\\DS224\home\claude\claude ai\Deploiement-Cabinet\Deployer le cabinet.lnk`
2. Au cabinet, lundi : copier ce `.lnk` sur le Bureau du PC médecin (AX8MAX),
   fermer Word et Excel, double-cliquer.
3. Le raccourci télécharge toujours la DERNIÈRE version depuis GitHub au
   moment du clic — inutile de le régénérer après une future correction.
4. Il installe automatiquement le poste médecin (racine `\\RDC\CabinetCardio`)
   et tente une installation à distance sur RDC (secrétariat). Si cette
   tentative échoue (réseau/pare-feu du cabinet), le bilan final l'indique
   clairement et demande de double-cliquer une fois, depuis RDC lui-même,
   sur `\\RDC\CabinetCardio\_Installation\installer_secretariat.cmd`.

Une fois ce déploiement fait et vérifié, effacer cette section (ou la
remplacer par le prochain point en cours).

## Le projet

Application de gestion de cabinet de cardiologie libérale (Beaumont-sur-Oise),
en VBA pour Word (`Cabinet.dotm`, poste médecin) et Excel (`Cabinet.xlsm`,
poste secrétariat), avec feuilles de soins papier Cerfa S3110 exclusivement
(pas de télétransmission).

- Sources : `Src/` (modules `.bas`/`.vba`, formulaires en JSON), assemblées
  par `Src/manifest.json` via un script de construction sur le NAS DS224
  (`CabinetCardio-Dev/Build/build.ps1`, ouvre Word et Excel en COM).
- Configuration et données de référence (non nominatives) : `Donnees/Config/`.
  **Les fichiers nominatifs (Patients.xlsx, Journal_*.xlsx, Agenda_*.xlsx)
  ne sont PAS dans ce dépôt** (purgés de l'historique pour raisons RGPD) :
  ne jamais les y remettre.
- Outils d'exploitation dans `Outils/` :
  - `maj_poste.ps1` : récupère GitHub sur le NAS, reconstruit les modèles,
    installe LE POSTE COURANT (rôle auto-détecté ou explicite).
  - `deployer_cabinet.ps1` : même chose + tente en plus l'installation à
    distance du poste secrétariat (RDC) depuis le poste médecin.
  - `creer_raccourci_deploiement.ps1` : dépose sur le NAS un raccourci
    auto-actualisé vers `deployer_cabinet.ps1`.
  - `verifier_nas.ps1` : contrôle en lecture seule que le NAS est à jour
    et intégralement déployable.
- Branche de travail : `claude/suivi-dev-logiciel-cabinet-fdjpa9`.
- `RECETTE_AUDIT.md` : protocole d'essai à blanc (bac à sable
  `C:\CabinetCardio-TEST`) des corrections issues de l'audit du 05/09/2026,
  décrites dans `OPTIMISATIONS_20260905.md` — à exécuter avant tout usage
  réel des nouveautés (circuit financier, agenda, Cerfa).

## Décisions à respecter

- Pas de PR sans demande explicite de l'utilisateur.
- Racine de données jamais écrasée par les scripts de mise à jour
  (robocopy additif sur `Donnees/`, jamais `/MIR`).
- Toute base nominative reste sur le NAS/les postes du cabinet, jamais
  sur GitHub.
