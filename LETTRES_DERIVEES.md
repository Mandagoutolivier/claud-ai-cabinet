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
   formule "Confraternellement". Le choix vient de la fiche correspondant
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
| `{formule}` | Bien cordialement / Confraternellement |
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

Confraternellement
{signature}
```

Le gabarit hospitalisation est au vouvoiement par défaut car il s'adresse à
un service ; le tutoiement s'applique seulement si le correspondant nommé
est marqué tutoiement. `{urgence}` : "programmée", "rapide", "en urgence".
`{verbe_hosp}` : "prendre en charge ce patient" / "cette patiente".
`{antecedents}` : ligne reprenant les antécédents cardiologiques majeurs.

## 5. Points à trancher par le médecin

- Confirmer "Cher ami" pour certains correspondants, ou "Cher confrère"
  pour tous.
- Confirmer la formule au vouvoiement : "Confraternellement" ou
  "Bien confraternellement".
- Fournir la liste des correspondants habituels par type de demande, pour
  pré-remplir le destinataire (centre de scintigraphie, radiologue, CCN).

## 6. Mise en œuvre

La correction se fait dans le module VBA de génération des courriers de
`Cabinet.dotm` et dans les modèles associés. Ces sources ne sont pas encore
dans le dépôt ; elles se trouvent dans le dossier `Src` et dans
`Donnees\Modeles` du paquet sur le NAS. Chaque type ci-dessus doit devenir un
gabarit distinct, alimenté par les variables de la section 3, et non un
découpage du compte rendu.
