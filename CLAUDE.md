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
- Config à retoucher À LA MAIN sur les deux postes (`Config\config.ini`,
  jamais écrasée) : `[COURRIER] PolitesseAuto=0` (décision 08/09 : appel et
  politesse dictés, plus de pré-remplissage ; `AppelAuto=0` est ajouté
  automatiquement).
- ECG (décision 10/09) : « Arrivé » sur ACCUEIL doit suffire (GDT écrit
  dans `[ECG] DossierGdt` = `\\AX8_MAX\Mandagout`, à vérifier dans le
  config.ini d'ACCUEIL) ; sexe (3110) et DDN dans le GDT ; veilleur
  `Build/ecg_valider_fenetre.ps1` sur AX8_Max (dossier Démarrage) qui envoie
  `[ECG] TouchesValidation` à la fenêtre `FenetreTitre` de Resting12Lead.
  Pont SQL (11/09) : le même veilleur écrit chaque nouvel IMPORT.GDT dans
  la table `patinfo` (base `ecgcenter`, login DMSNIS de `SetSQLInfo.ini`)
  sur SQL Server Express local d'AX8_Max, préparée une fois par
  `Build/installer_ecg_sql.ps1` (admin) ; activer par `[ECG] SqlActif=1`
  et cocher « Connection Système Info DMS » dans Resting12Lead.
- Lettres dérivées (décision 10/09) : architecture R12 = UN appel API à la
  correction rend courrier + blocs `DEMANDE_DESTINATION` (`modDemandesR12`,
  prompt `Config\prompts\demandes_r12.txt`, `[DERIVEES] Mode=R12`) ; le VBA
  assemble les lettres à la finalisation. `Mode=Profils` = ancien mode.
- Données sur le NAS (décision 12/09) : la racine doit être un dossier
  partagé du DS224 (`\\DS224\CabinetCardio`, RAID + Hyper Backup vers le
  NAS du domicile), plus jamais sur un PC. Migration par le
  raccourci « Deployer le cabinet » (ACCUEIL allumé, Word/Excel fermés :
  il propose la copie vérifiée puis pose `RACINE_DEPLACEE.txt` dans
  l'ancienne racine, que `modConfig.Racine` suit automatiquement), ou par
  `Outils/migrer_racine_nas.ps1` ;
  `deployer_cabinet.ps1` détecte `$RacineNas\Base` et bascule les deux
  postes dessus (paquet d'installation déposé dans `\\DS224\CabinetCardio\_Installation`).
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
- Scripts de construction et d'installation dans `Build/` (`build.ps1`,
  `installer_cabinet.ps1`, `installer_relais.ps1`, `sync_startup.ps1`,
  `init_donnees.ps1`) : versionnés ici, publiés vers le NAS par les outils.
- Raccourcis Word : Ctrl+Alt+N/D/P/G/V/B, F6 (patient), **Ctrl+Alt+Maj+C**
  correction (Ctrl+Alt+C est réservé à l'ancien complément ChatGPT),
  **Ctrl+Alt+Maj+A / B / D** = boutons PowerMic : destinataire, appel,
  finaliser (corriger + demandes dans le même fichier + `[SORTIE] Dossier`
  + secrétariat). Dragon appelle plutôt les macros de `modPowerMic`
  (`wd.Run "Cabinet_A_NouvelleLettre"`, `_B_FormuleAppel`, `_D_Finaliser`,
  `_P_Patient`).
- Outils d'exploitation dans `Outils/` :
  - `maj_poste.ps1` : récupère GitHub sur le NAS, reconstruit les modèles,
    installe LE POSTE COURANT (rôle auto-détecté ou explicite).
  - `deployer_cabinet.ps1` : même chose + tente en plus l'installation à
    distance du poste secrétariat (RDC) depuis le poste médecin.
  - `creer_raccourci_deploiement.ps1` : dépose sur le NAS un raccourci
    auto-actualisé vers `deployer_cabinet.ps1`.
  - `migrer_racine_nas.ps1` : déplace la racine des données vers le NAS
    et ré-enregistre le poste (copie vérifiée, ancien dossier renommé).
  - `creer_raccourci_domicile.ps1` : raccourci Bureau du PC domicile qui
    lance la dernière `maj_poste.ps1 -Role Tous` (médecin + secrétariat).
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
