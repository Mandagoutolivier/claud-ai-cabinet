# Optimisations suite à l'audit du 05/09/2026

Corrections apportées d'après `Analyse_modele_gestion_cabinet_20260905.md`
(priorités hautes et moyennes).

## 1. Circuit financier — plus d'honoraires fantômes

- `modBaseIO.FermerTransaction` : l'échec de `Workbook.Save` était avalé.
  Il lève désormais l'erreur `vbObjectError + 603` ; **rien n'est
  silencieusement perdu**, la secrétaire est prévenue et recommence.
- `modBaseIO` : ajout de `ModifierLignes`, `TrouverLignes` et
  `AssurerColonnes` (migration des colonnes des classeurs existants).
- `modJournal` : colonnes explicites — `DateActe` / `DateSaisie` distinctes,
  `SeanceID`, `ConsultationID`, `CodeOfficiel`, `MontantDu` /
  `MontantRegle`, `DateEncaissement`, `FeuilleSoinsEtat` /
  `DateFeuilleSoins`. `Paye` n'est mis à « O » que si l'argent est reçu.
- `modActes.EnregistrerSeance` : séance **idempotente**. Une consultation
  déjà portée au journal lève `vbObjectError + 610` au lieu de créer une
  seconde fois les mêmes honoraires (réimpression et règlement se font par
  les boutons dédiés). Répartition en cascade du montant réglé.
- `modActes.EnregistrerReglement` : encaissement différé (virement, chèque
  reçu plus tard), renvoie le solde restant dû.
- `modActes.VerifierCotation` + `Config/regles_cotation.txt` : garde-fou
  paramétrable par le cabinet (`INTERDIT=`, `UNIQUE=`, `MAXPARSEANCE=`).
  Ce n'est pas la NGAP : la cotation reste sous la responsabilité du médecin.

## 2. Feuille de soins (cerfa S3110, papier)

- `modCerfaPrint.ImprimerFeuille` refuse explicitement au-delà de
  `[CERFA] LignesMax` (4) au lieu de tronquer en silence.
- Chaque ligne porte sa **date réelle d'acte** (`vbObjectError + 702` si
  elle manque).
- Rubriques ajoutées : `PATIENT_NOM/DDN/NIR` distinctes de l'assuré quand
  le patient n'ouvre pas droit, et `MEDECIN_TRAITANT` (parcours de soins).
- `ufChoixActe` : la case « imprimée » ne prouve pas la sortie papier — la
  secrétaire confirme (`Remise` / `EchecImpression`) ; boutons
  « Document seul », « Enregistrer un règlement », « Réimprimer ».

## 3. Agenda — chevauchements

- `modAgenda.ConflitRdv` : détection réelle des **chevauchements**
  (début + durée), et non plus de la seule égalité d'horaire.
- `modAgenda.AjouterRdv` valide date, heure et durée (erreurs 620 à 624) et
  refuse un créneau occupé ; `forcer:=True` autorise un double créneau assumé.
- `ufRdvEdit` : validation stricte et demande de confirmation explicite en
  cas de chevauchement. `modAgendaVue.CreneauLibre` s'appuie sur la même règle.

### 3 bis. Agenda — second lot (06/09/2026)

- **Correctif bloquant** : `ConflitRdv` passait des heures `hh:mm` à
  `IntervallesSeChevauchent` qui attend des minutes → erreur 13 à chaque
  prise de RDV. Converti par `MinutesDepuisMinuit`.
- **Navigation par mois** : boutons « << Mois » / « Mois >> » (en vue
  semaine : même jour de semaine un mois plus tôt/tard), « Aller à une
  date… » (`jj/mm/aaaa` ou `+N`/`-N` jours).
- **Vue mois** : calendrier du mois, nombre de RDV et indisponibilités par
  jour, double-clic sur un jour → sa semaine. Bouton « Vue mois / semaine ».
- **Déplacer / modifier un RDV** : depuis le double-clic sur un RDV, même
  écran que la prise de RDV ; garde l'ID (nouvelle ligne seulement si
  l'année change) ; le RDV déplacé ne se gêne pas lui-même.
- **Indisponibilités** (`Statut = Bloque`, `TypeActe = INDISPO`) : congés,
  formation… sur une période, journée entière ou plage horaire ; grisées
  dans l'agenda, refusées à la prise de RDV (avec le motif), levables ;
  exclues de la liste d'arrivée des patients.
- **Lien RDV ↔ consultation** : à l'enregistrement de la séance, le RDV du
  patient ce jour-là passe à « Honore » et reçoit `ConsultationID` ;
  le journal porte `RdvID`. Colonnes ajoutées par migration automatique.
- **Rafraîchissement multi-poste** : `[AGENDA] RafraichirSecondes` (60) —
  redessin silencieux uniquement si le fichier agenda a changé.
- **Patient provisoire** : bouton « Nouveau patient (nom + téléphone) »
  dans la prise de RDV ; fiche minimale marquée `Provisoire = O`, affichée
  « [prov.] », la marque tombe quand la date de naissance est saisie.
- **Feuille du jour imprimable** : bouton « Imprimer le jour » — aperçu
  paysage + PDF `Echange\Jour_AAAA-MM-JJ.pdf` pour le poste médecin.
- **Couleur par motif** : `[AGENDA] CouleurParActe=ETT:C6E0B4;…` pour les
  RDV prévus ; le statut garde priorité (arrivé, absent, vu, indisponible).
- **Durée par motif** : `[AGENDA] DureeParActe=CS:15;APC:30;…` — la durée
  suit le motif choisi (valeurs par défaut à ajuster).

## 3 ter. Honoraires et encaissements (06/09/2026, `modHonoraires`)

Bouton séparé à l'accueil, feuille « Honoraires ». Lecture du journal ;
seules écritures : règlement (`EnregistrerReglement`), `RelanceLe`,
`ChequeRemisLe`. Aucune ligne d'acte n'est créée par ce module.

- **Impayés** : séances non soldées, ancienneté, rouge au-delà de
  `[HONORAIRES] RelanceApresJours` sans relance, jaune si virement attendu
  ou tiers payant ; actions règlement / relance / note / fiche.
- **Note d'honoraires** : n° = séance, actes, montants, « Acquittée le »
  ou « Reste dû », PDF `Actes\Notes\`.
- **Chèques** : liste des chèques reçus non remis, bordereau PDF
  `Actes\Remises\`, marquage « remis le » après confirmation.
- **Recettes du mois** : encaissements réels sans identité de patient,
  totaux par mode et par jour, PDF + Excel `Actes\Recettes\` (comptable).
- **Tableau de bord** : facturé / encaissé / reste dû par mois, par acte.

Limite connue : une même ligne d'acte réglée en deux fois ne garde que la
date du dernier encaissement (le journal est tenu par ligne d'acte, pas par
règlement). Un registre des règlements est l'évolution suivante si le
comptable le demande.

## 4. Dates et identités

- `modTexte` : `DateFrValide` (calendrier réel, années bissextiles, pas de
  date future pour une naissance), `DateFr`, `HeureValide`,
  `MinutesDepuisMinuit`, `IntervallesSeChevauchent`.
- `ufPatientEdit` délègue à `modTexte.DateFrValide` (32/13/2026 refusé).
- `modPatient.DossierPatient` : le dossier est nommé à partir de
  **l'identifiant stable** (`ID_Nom_Prenom`). Un changement d'état civil
  renomme le dossier au lieu d'en créer un second ; les anciens dossiers
  `Nom_Prenom_ID` sont migrés automatiquement.

## 5. Courriers

- `modValidation` : identifiant de document **stable** (`ConsultationID`,
  posé à la première validation et conservé dans le document), `DateActe`
  figée, transmis dans le drapeau. Les corrections successives produisent
  des versions (`... v2.docx`) au lieu d'écraser ou de créer un doublon.
- `modDerivees` : le **motif saisi librement** par le médecin est
  désormais balisé et scanné comme le reste avant tout envoi à l'API.
- `modAnonymise.VerifierBalisesRetour` vérifie aussi qu'aucune balise
  présente dans le texte envoyé n'a **disparu** de la réponse (une balise
  perdue = une identité non réinjectée).
- `prompts/derivee.txt` : suppression de la contradiction sur la dernière
  phrase (la lettre se termine par la demande, pas par une formule de
  courtoisie) ; l'exemple d'arrêt de traitement à 48 heures est explicitement
  une illustration de style, jamais une règle médicale à inventer.

## 6. Configuration

- `Debug=0` par défaut (le mode debug écrit le contenu envoyé à l'API).
- `[DERIVEES] AutoValider=0` : une demande d'examen engage le médecin, elle
  est produite en brouillon et relue avant transmission.
- Section `[ECG]` rétablie dans la configuration déployée.
- `[CERFA] LignesMax=4`.
- Nouveau `Config/regles_cotation.txt` (et sa copie `Src/ConfigDefaut`).
