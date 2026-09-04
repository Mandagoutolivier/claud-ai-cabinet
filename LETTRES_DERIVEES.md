# Lettres dérivées : cahier des charges et gabarits

Ce document définit les courriers que le logiciel doit produire à partir d'une
consultation, en plus du compte rendu lui-même. Il remplace la logique actuelle
qui découpe le compte rendu de consultation : une lettre dérivée est une
**demande** adressée à un confrère, pas un résumé.

## 1. Périmètre

### Pas de courrier dérivé

Les examens réalisés au cabinet par le Dr Mandagout ne donnent lieu à aucune
lettre : échocardiographie, Holter ECG, MAPA (mesure ambulatoire de la
pression artérielle).

### Courriers dérivés à produire

| Code | Type de demande | Destinataire habituel |
|------|-----------------|-----------------------|
| `TE`  | Épreuve d'effort | Cardiologue, centre d'explorations |
| `SCINTI` | Scintigraphie myocardique, selon modalité (voir 4.2) | Médecine nucléaire |
| `SCANCORO` | Scanner coronaire, avec ou sans score calcique | Radiologie |
| `IRM` | IRM cardiaque de repos ou de stress (voir 4.4) | Radiologie / centre d'IRM |
| `AVIS` | Avis spécialisé : neurologie, gastro-entérologie, SAOS (pneumologie / sommeil), autre spécialité saisie librement | Spécialiste concerné |
| `HOSP` | Demande d'hospitalisation, notamment au CCN (Centre Cardiologique du Nord) | Service hospitalier |

## 2. Règles communes

1. **Registre confraternel cohérent.** Appel "Cher confrère" ou "Chère consœur"
   selon le sexe du correspondant ; "Cher ami" / "Chère amie" si le
   correspondant est marqué comme proche dans la base. Tutoiement par défaut
   avec ces appels. Formule finale : "Bien cordialement". Jamais de
   "Madame, Monsieur" ni de "Veuillez agréer" dans une lettre à un confrère.
2. **Vouvoiement possible** pour un correspondant inconnu ou un service
   hospitalier : appel "Cher confrère" / "Chère consœur", vouvoiement,
   formule "Bien confraternellement". Le choix vient de la fiche correspondant
   (champ `tutoiement` oui/non), à défaut vouvoiement.
3. **La demande ouvre la lettre.** La première phrase dit ce qui est demandé
   et pour qui : "Merci de réaliser [examen] à [Madame/Monsieur Prénom NOM],
   [âge] ans."
4. **Deux à quatre lignes cliniques au maximum**, choisies pour la demande :
   symptômes ou leur absence, ECG, facteurs de risque ou antécédent utile,
   traitement pertinent. Pas de "Je revois ... pour un contrôle", pas de
   "Au total", pas de "Je te tiens informé" : ces phrases appartiennent au
   compte rendu, pas à la demande.
5. **Aucune contradiction** entre le motif et la demande : on ne qualifie pas
   le bilan de "rassurant" dans une lettre qui demande de le compléter. Le
   logiciel n'insère la conclusion du compte rendu que si elle a été
   explicitement cochée par le médecin.
6. **Une seule phrase de courtoisie**, propre à chaque gabarit, juste avant
   la formule finale.
7. **Signature** : "Docteur Olivier MANDAGOUT" suivi des mentions habituelles
   du cabinet.
8. **Objet** en tête de lettre, pour le secrétariat du destinataire :
   "Objet : demande de [examen] – [NOM Prénom], né(e) le [date]".

## 3. Variables disponibles

| Variable | Contenu |
|----------|---------|
| `{civilite}` | Madame / Monsieur |
| `{prenom}` `{nom}` | Prénom, NOM en majuscules |
| `{age}` | Âge en années |
| `{ddn}` | Date de naissance |
| `{appel}` | Cher confrère / Chère consœur / Cher ami / Chère amie |
| `{tu_vous}` | Sélecteur de conjugaison, tutoiement ou vouvoiement |
| `{formule}` | Bien cordialement / Bien confraternellement |
| `{symptomes}` | Ligne clinique, par ex. "Elle est totalement asymptomatique." |
| `{ecg}` | Ligne ECG, par ex. "L'électrocardiogramme est normal." |
| `{fdr}` | Facteurs de risque, facultatif |
| `{motif}` | Motif libre saisi au moment de la demande |
| `{traitement}` | Traitement en cours, facultatif |
| `{signature}` | Bloc signature |

Les lignes `{symptomes}`, `{ecg}`, `{fdr}`, `{traitement}` sont omises si vides,
sans laisser de ligne blanche.

## 4. Gabarits

Les gabarits ci-dessous sont écrits au tutoiement. En vouvoiement, le logiciel
conjugue "Je te serais reconnaissant" en "Je vous serais reconnaissant", et
remplace la formule finale.

### 4.1 Épreuve d'effort (`TE`)

```
Objet : demande d'épreuve d'effort – {nom} {prenom}, né(e) le {ddn}

{appel},

Merci de réaliser une épreuve d'effort à {civilite} {prenom} {nom}, {age} ans.
{symptomes}
{ecg}
{fdr}
{motif}

Je te serais reconnaissant de bien vouloir réaliser cet examen afin de
compléter ce bilan.

{formule}
{signature}
```

Exemple attendu :

```
Cher confrère,

Merci de réaliser une épreuve d'effort à Madame Yamna FIZAINE, 75 ans.
Elle est totalement asymptomatique.
L'électrocardiogramme est normal.

Je te serais reconnaissant de bien vouloir réaliser cet examen afin de
compléter ce bilan.

Bien cordialement
Docteur Olivier MANDAGOUT
```

### 4.2 Scintigraphie myocardique (`SCINTI`)

Modalités à proposer au moment de la demande, une seule cochée :

| Modalité | Texte inséré dans la demande |
|----------|------------------------------|
| Effort | "une scintigraphie myocardique de perfusion d'effort" |
| Pharmacologique | "une scintigraphie myocardique de perfusion sous stress pharmacologique (dipyridamole ou régadénoson)" |
| Couplée | "une scintigraphie myocardique de perfusion sous effort couplé à un stress pharmacologique" |
| Repos seul | "une scintigraphie myocardique de perfusion de repos" |

```
Objet : demande de scintigraphie myocardique – {nom} {prenom}, né(e) le {ddn}

{appel},

Merci de réaliser {modalite_scinti} à {civilite} {prenom} {nom}, {age} ans.
{symptomes}
{ecg}
{fdr}
{traitement}
{motif}

Je te serais reconnaissant de bien vouloir réaliser cet examen et de
m'adresser le compte rendu.

{formule}
{signature}
```

Le logiciel rappelle au médecin, au moment de la demande, de préciser dans
`{motif}` une contre-indication à l'effort ou un bloc de branche gauche
lorsqu'ils justifient la modalité pharmacologique.

### 4.3 Scanner coronaire (`SCANCORO`)

Options : "avec score calcique" (coché par défaut), "score calcique seul".

```
Objet : demande de scanner coronaire – {nom} {prenom}, né(e) le {ddn}

{appel},

Merci de réaliser un scanner coronaire{option_scan} à {civilite} {prenom} {nom}, {age} ans.
{symptomes}
{ecg}
{fdr}
{motif}

La fonction rénale est {fonction_renale}.

Je te serais reconnaissant de bien vouloir réaliser cet examen.

{formule}
{signature}
```

`{option_scan}` vaut " avec score calcique" ou " (score calcique seul)".
`{fonction_renale}` est saisie ou reprise de la dernière créatinine connue ;
la ligne est omise si l'information manque.

### 4.4 IRM cardiaque (`IRM`)

Modalités : repos ; stress sous dobutamine ; stress sous vasodilatateur
(adénosine ou régadénoson).

```
Objet : demande d'IRM cardiaque – {nom} {prenom}, né(e) le {ddn}

{appel},

Merci de réaliser une IRM cardiaque {modalite_irm} à {civilite} {prenom} {nom}, {age} ans.
{symptomes}
{ecg}
{echo}
{motif}

{question_posee}

Je te serais reconnaissant de bien vouloir réaliser cet examen.

{formule}
{signature}
```

`{modalite_irm}` : "de repos", "de stress sous dobutamine", "de stress sous
vasodilatateur". `{echo}` : une ligne reprenant le résultat échographique
utile, par exemple "La FEVG est estimée à 45 %." `{question_posee}` : phrase
libre, par exemple "La question posée est celle d'une viabilité myocardique."

### 4.5 Avis spécialisé (`AVIS`)

Spécialité choisie dans une liste : neurologie, gastro-entérologie, SAOS
(pneumologie / médecine du sommeil), puis "autre" avec saisie libre.

```
Objet : demande d'avis {specialite} – {nom} {prenom}, né(e) le {ddn}

{appel},

Je te confie {civilite} {prenom} {nom}, {age} ans, pour un avis {specialite}.
{motif}
{symptomes}
{traitement}

{question_posee}

Je te serais reconnaissant de bien vouloir {verbe_avis}.

{formule}
{signature}
```

`{specialite}` : "neurologique", "gastro-entérologique", "pneumologique à la
recherche d'un syndrome d'apnées du sommeil", ou le libellé saisi.
`{verbe_avis}` : "recevoir ce patient" / "recevoir cette patiente" selon le
sexe.

### 4.6 Demande d'hospitalisation (`HOSP`)

Établissement choisi dans la base des correspondants, CCN proposé en premier.
Degré d'urgence : programmée, rapide (sous 48 h), urgente (jour même).

```
Objet : demande d'hospitalisation {urgence} – {nom} {prenom}, né(e) le {ddn}

{appel},

Je vous adresse {civilite} {prenom} {nom}, {age} ans, pour une hospitalisation
{urgence} dans votre service.
{motif}
{symptomes}
{ecg}
{echo}
{traitement}
{antecedents}

{question_posee}

Je vous remercie de bien vouloir {verbe_hosp}.

Bien confraternellement
{signature}
```

Le gabarit hospitalisation est au vouvoiement par défaut car il s'adresse à
un service ; le tutoiement s'applique seulement si le correspondant nommé
est marqué tutoiement. `{urgence}` : "programmée", "rapide", "en urgence".
`{verbe_hosp}` : "prendre en charge ce patient" / "cette patiente".
`{antecedents}` : ligne reprenant les antécédents cardiologiques majeurs.

## 5. Décisions du médecin (04/09/2026)

- **Appel** : "Cher confrère" / "Chère consœur" par défaut ; "Cher ami" /
  "Chère amie" pour les correspondants marqués comme proches dans la base des
  correspondants (champ `proche` oui/non). Un correspondant proche est
  tutoyé.
- **Formule au vouvoiement** : "Bien confraternellement". Au tutoiement :
  "Bien cordialement".
- **Destinataire pré-rempli** : la liste des correspondants habituels par type
  de demande est définie dans le classeur Excel des correspondants du paquet
  (`Donnees`). Le logiciel y lit, pour chaque code de demande (`TE`,
  `SCINTI`, `SCANCORO`, `IRM`, `AVIS` par spécialité, `HOSP`), le ou les
  correspondants à proposer, le premier étant sélectionné par défaut.

## 6. Mise en œuvre (04/09/2026, canevas R12 intégrés)

### Déclenchement automatique à la validation

À la validation du courrier principal (Ctrl+Alt+V, après correction), le
logiciel repère les demandes d'examen ou d'avis et génère les lettres sans
intervention. Un courrier qui est lui-même une demande n'en engendre pas.

1. **Repérage** dans `Config\demandes\`, trois fichiers texte modifiables au
   Bloc-notes par le médecin :
   - `declencheurs.txt` : formules exprimant une demande, une par ligne,
     `*` pour un texte variable (« je prescris* », « je l'adresse*pour* »).
   - `examens.txt` : `expression|CODE_DU_PROFIL` (« test d'effort|TEST_EFFORT »).
   - `exclusions.txt` : formules qui annulent (« avait réalisé* », « pourrait réaliser* »).
   Une demande est reconnue quand un déclencheur et un examen sont dans la
   même phrase sans exclusion. Comparaison sans casse ni accents, en mots
   entiers, l'expression la plus longue l'emporte (« irm de stress » avant
   « irm », « coroscanner » ne déclenche pas « scanner »).
2. **Profil** `Config\profils\CODE.ini` (33 canevas issus de R12, UTF-8,
   modifiables) : premier paragraphe imposé, rubriques dans l'ordre avec
   leur statut (obligatoire, si présente, optionnelle, interdite), préfixes
   de phrase, mots-clés de sélection, longueur (COURTE, MOYENNE, DETAILLEE),
   objectif, modalités, `TYPE_BASE` (codes du classeur des spécialistes).
   Profil absent : `AUTRE_EXAMEN.ini`.
3. **Destinataire** : spécialiste de priorité 1 pour `TYPE_BASE` dans
   `Specialistes_ParType` ; à défaut un bloc « DESTINATAIRE À COMPLÉTER ».
4. **Rédaction** par l'API : courrier source anonymisé, identité du patient
   balisée, phrases de prescription (qui portent l'examen précis, sa
   modalité et le motif), consignes du profil. L'argumentaire se limite au
   motif, à l'examen clinique, à l'ECG et au traitement, plus ou moins
   développés selon la longueur du profil.
5. **Validation** : `AutoValider=1` enregistre docx et PDF dans le dossier du
   patient et transmet au secrétariat ; `AutoValider=0` laisse les lettres
   ouvertes à relire. Le message de validation liste les lettres produites.

### Commande manuelle Ctrl+Alt+D

Choix du profil (liste filtrable), de la modalité si le profil en a, motif
facultatif, destinataire pré-rempli par type et trié par priorité
(« Nouveau... » = liste générale), puis génération.

### Fichiers

| Élément | Fichier |
|---------|---------|
| Repérage, profils, consignes API, génération automatique | `Src/Word/modDemandes.bas` |
| Commande manuelle, destinataires, appel API, création du courrier | `Src/Word/modDerivees.bas` |
| Validation réutilisable et appel des demandes automatiques | `Src/Word/modValidation.bas` (`ValiderDocument`) |
| Appel et formule par défaut selon tutoiement, bloc destinataire en gras à l'interligne du corps | `Src/Word/modCourrier.bas` |
| Prompt des dérivées sans courriers de référence | `Src/Word/modClaude.bas` (`ChargerPrompt`, `avecReferences`) |
| Gabarit du prompt | `Donnees/Config/prompts/derivee.txt` |
| Dictionnaires de repérage | `Donnees/Config/demandes/*.txt` |
| Canevas | `Donnees/Config/profils/*.ini`, `CATALOGUE.txt` |
| Classeur des spécialistes | `Donnees/Base/Correspondants_Specialistes.xlsx` |
| Configuration | `config.ini` `[DERIVEES]` (Automatique, AutoValider, DossierProfils, DossierDemandes, Types, FichierSpecialistes) et `[COURRIER]` (formules par défaut) |
| Test hors ligne | `Src/Word/modTests_Word.bas`, `Test_DERIVEES` |

### Correspondance profils et codes du classeur des spécialistes

`TYPE_BASE` des profils a été aligné sur les codes `TypeExamen` du classeur :
`AVIS_NEUROLOGIQUE.ini` → `AVIS_NEUROLOGIE`, `AVIS_GASTROENTEROLOGIQUE.ini` →
`AVIS_GASTRO_ENTEROLOGIE`, `RECHERCHE_SAOS.ini` → `SAOS,AVIS_PNEUMOLOGIE`,
`IRM_CARDIAQUE.ini` et `IRM_DE_STRESS.ini` → `IRM`, `HOSPITALISATION_CCN.ini` →
`HOSPITALISATION_CCN`, `HOSPITALISATION_AUTRE_CENTRE.ini` → `HOSPITALISATION`.
Le classeur ne contient pas encore de ligne `HOSPITALISATION_CCN`,
`HOSPITALISATION` ni `AVIS_NEUROLOGIE` : les ajouter dans
`Specialistes_ParType` (CCN en priorité 1) pour que le destinataire soit
pré-rempli, sinon la lettre sort avec un bloc « à compléter ».

### Ordre des formules d'appel et de politesse

1. Colonnes `FormuleAppel` / `FormulePolitesse` de la ligne `Specialistes_ParType`.
2. À défaut, `FormuleAppel` de la fiche `Specialistes`.
3. À défaut, `config.ini` `[COURRIER]` : `AppelProche` / `PolitesseProche` si
   `TutoiementVouvoiement` vaut `tu`, sinon `AppelDefaut` / `PolitesseDefaut`.

Le tutoiement pilote aussi la rédaction du corps par l'API.

### Mise en page de l'adresse du correspondant

Sur tous les courriers, le bloc adresse du correspondant est un seul
paragraphe à sauts de ligne manuels, à l'interligne du corps
(`[COURRIER] Interligne`), sans espacement entre les lignes, entièrement en
gras (`modCourrier.MettreEnFormeDestinataire`).
