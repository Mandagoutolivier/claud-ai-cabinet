# Cabinet Cardio — logiciel de gestion (Dr Mandagout)

## 📌 Pense-bête en cours

**Déploiement du 07/09/2026 — état :**

- Poste médecin (AX8_Max) : déployé par le raccourci « Deployer le cabinet »
  (version f06572f), racine `\\ACCUEIL\CabinetCardio`, relais Ctrl+Alt+…
  testé OK, clé API présente. Tâche planifiée refusée (sans conséquence).
- Poste secrétariat ACCUEIL : installé le 07/09 (partage créé), données de
  l'ancien poste RDC rapatriées (agenda, journal, dossiers). **Reste à faire
  une fois** : double-cliquer, depuis ACCUEIL, Word/Excel fermés,
  `\\ACCUEIL\CabinetCardio\_Installation\installer_secretariat.cmd`
  pour y mettre le dernier `Cabinet.xlsm` (agenda mois, honoraires).
  L'installation à distance (WinRM) ne passe pas au cabinet : c'est normal.
- Ancien poste RDC : à neutraliser (renommer `C:\CabinetCardio`) pour éviter
  deux agendas.
- Ensuite : recette `RECETTE_AUDIT.md` (chapitres 2, 4, 5 bis) avant usage
  réel du circuit financier.

Mettre à jour cette section au fil de l'eau ; l'effacer quand tout est vérifié.

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
