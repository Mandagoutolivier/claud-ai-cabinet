# Recette avant utilisation en conditions réelles

Les corrections issues de l'audit du 05/09/2026 (circuit financier, feuille
de soins, agenda, dates) n'ont jamais tourné sur des cas réels. Cette fiche
sert à les éprouver **sur des données de test**, pas sur une vraie journée.

## 0. Se placer en bac à sable

Ne faites JAMAIS cette recette sur `C:\CabinetCardio` du secrétariat.

```powershell
# racine de test, séparée des vraies données
$test = "C:\CabinetCardio-TEST"
New-Item -ItemType Directory -Force -Path $test | Out-Null
powershell -NoProfile -ExecutionPolicy Bypass -File C:\Temp\maj_poste.ps1 `
    -Racine $test -Role Secretaire -BaseInitiale ""
```

Le logiciel crée alors une base vide (`Patients.xlsx` avec ses colonnes).
Créez deux ou trois patients fictifs, un correspondant fictif, et travaillez
uniquement avec eux.

Pour revenir aux vraies données : `notepad %APPDATA%\CabinetCardio\chemin.txt`
et remettre `C:\CabinetCardio`.

## 1. Tests automatiques

Word : Alt+F8 → `Test_TOUT` (ou les `Test_*` de `modTests_Word`).
Excel : Alt+F8 → les `Test_*` de `modTests_Excel`.
Résultat attendu : aucune ligne en échec dans `Logs\`.

## 2. Circuit financier (le point le plus sensible)

| # | Ce qu'on fait | Attendu |
|---|---|---|
| 2.1 | Valider un courrier, enregistrer la séance (CSC, CB, montant réglé = total) | 2 lignes au journal (CSC + DEQP003), `Paye` = O, `DateActe` = date réelle |
| 2.2 | **Rejouer** le même document (revalider, réenregistrer) | REFUS explicite « déjà enregistrée au journal » — aucun honoraire en double |
| 2.3 | Corriger le courrier puis le revalider | nouvelle version `... v2.docx`, même `ConsultationID`, toujours une seule séance |
| 2.4 | Séance avec montant réglé = 0 (virement attendu) | `Paye` = N, `MontantDu` renseigné, `MontantRegle` = 0 |
| 2.5 | « Enregistrer un règlement » pour cette séance | solde mis à jour, `DateEncaissement` renseignée, `Paye` = O quand tout est payé |
| 2.6 | « Document seul (aucun acte) » | document classé, AUCUNE ligne au journal |
| 2.7 | Débrancher le dossier partagé pendant l'enregistrement | message d'échec explicite ; vérifier qu'aucune ligne partielle n'est restée |

Le 2.7 est le plus important : c'est le bug qui perdait silencieusement des
enregistrements. Testez-le vraiment (renommez le dossier `Base` une seconde).

## 3. Feuille de soins (papier, cerfa S3110)

| # | Ce qu'on fait | Attendu |
|---|---|---|
| 3.1 | Imprimer vers PDF une séance à 2 actes | dates réelles sur chaque ligne, total correct |
| 3.2 | Séance à 5 actes | REFUS explicite (4 lignes maximum) au lieu d'une troncature |
| 3.3 | Patient ≠ assuré (renseigner `AssureNom`) | rubriques patient ET assuré distinctes |
| 3.4 | Répondre « non » à « la feuille est-elle sortie ? » | état `EchecImpression` au journal, pas `Remise` |
| 3.5 | « Réimprimer la feuille » | même contenu, AUCUN nouvel honoraire |
| 3.6 | Superposer le PDF à une vraie liasse vierge | calage des positions ; ajuster `cerfa_positions.txt` si besoin |

## 4. Agenda

| # | Ce qu'on fait | Attendu |
|---|---|---|
| 4.1 | RDV 10:00 durée 30, puis RDV 10:15 | avertissement de chevauchement, confirmation demandée |
| 4.2 | Confirmer | les deux RDV existent (double créneau assumé) |
| 4.3 | RDV 10:30 après un 10:00/30 | aucun avertissement (pas de chevauchement) |
| 4.4 | Date 32/13/2026, heure 25:00 | refus de saisie |
| 4.5 | Annuler un RDV puis reprendre son créneau | plus d'avertissement |
| 4.6 | Double-clic sur un RDV → « Déplacer / modifier », changer l'heure | le RDV bouge, garde son ID, l'ancien créneau est libre |
| 4.7 | Déplacer un RDV sur un créneau occupé | avertissement de chevauchement, confirmation demandée |
| 4.8 | « Bloquer des créneaux » : deux jours entiers, motif « Congés » | jours grisés INDISPONIBLE, prise de RDV avertit « INDISPONIBLE : Congés » |
| 4.9 | Double-clic sur l'indisponibilité → « Lever » | créneau redevenu libre |
| 4.10 | « << Mois » / « Mois >> » depuis la semaine | même jour de semaine un mois avant/après |
| 4.11 | « Vue mois / semaine » puis double-clic sur un jour | calendrier du mois (nb de RDV par jour), puis semaine de ce jour |
| 4.12 | « Aller à une date… » avec `+90` | semaine dans trois mois |
| 4.13 | Enregistrer la séance d'un patient qui avait RDV ce jour | le RDV passe à « (vu) », colonne `ConsultationID` remplie, `RdvID` au journal |
| 4.14 | Sur l'autre poste, prendre un RDV ; attendre `RafraichirSecondes` | la grille se met à jour seule, sans perdre la sélection |
| 4.15 | Prise de RDV → « Nouveau patient (nom + téléphone) » | RDV créé, affiché « [prov.] » ; la fiche apparaît dans la base marquée provisoire |
| 4.16 | Ouvrir cette fiche, saisir la date de naissance, enregistrer | la marque provisoire disparaît, « [prov.] » aussi |
| 4.17 | « Imprimer le jour » | aperçu paysage (heure, patient, tél, motif, durée, notes, statut) + PDF `Echange\Jour_AAAA-MM-JJ.pdf` |
| 4.18 | RDV ETT, Holter, consultation prévus le même jour | trois couleurs différentes ; un RDV « arrivé » reste bleu quel que soit le motif |

## 5. Dates et dossiers

| # | Ce qu'on fait | Attendu |
|---|---|---|
| 5.1 | Naissance 30/02/1950 ou une date future | refus |
| 5.2 | Renommer un patient (mariage) puis créer un courrier | le dossier existant est RENOMMÉ, pas dupliqué |
| 5.3 | Rouvrir un courrier de la veille | la date affichée reste celle de sa création |

## 6. Lettres de demande

`[DERIVEES] AutoValider=0` : les demandes sont produites **en brouillon**.
Vérifier qu'à la validation d'un courrier contenant « je prescris une
scintigraphie… », la lettre est générée, ouverte, et **non transmise** au
secrétariat avant relecture.

## Que faire en cas d'échec

Notez le numéro de l'étape, le message exact et joignez le dernier fichier
de `Logs\`. Tant qu'un point du chapitre 2 échoue, ne basculez pas la
comptabilité sur cette version.
