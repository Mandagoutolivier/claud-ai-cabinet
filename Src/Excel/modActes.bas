Attribute VB_Name = "modActes"
Option Explicit
' =====================================================================
' modActes - Nomenclature des actes (Config\Nomenclature.xlsx, editable
' par la secretaire) et enregistrement d'une seance au journal.
'
' Colonnes de la nomenclature :
'   Code            code INTERNE du cabinet (CS, CSC, ETT, HOLTER, MAPA, APC)
'   CodeOfficiel    code porte sur la feuille de soins (ex DZQM006) ;
'                   a defaut, l'ancienne colonne LibelleCerfa est utilisee
'   LibelleCourt    libelle affiche a la secretaire
'   Tarif           honoraire demande (ce que paie le patient)
'   BaseRemboursement  base de remboursement si elle differe (information)
'   Remboursable    O/N (information ; N = acte non pris en charge)
'   CodeAssocie / TarifAssocie   acte associe genere avec le principal
'   Actif           0 = masque
'
' Les tarifs et les associations sont des PARAMETRES du cabinet : ils ne
' valent pas validation conventionnelle. Les cumuls sont controles par
' Config\regles_cotation.txt (voir VerifierCotation).
' =====================================================================

Public Function EntetesNomenclature() As Variant
    EntetesNomenclature = Array("Code", "CodeOfficiel", "LibelleCourt", "LibelleCerfa", _
                                "Tarif", "BaseRemboursement", "Remboursable", _
                                "CodeAssocie", "TarifAssocie", "Depassement", "Actif")
End Function

Public Function Nomenclature() As Collection
    Dim tous As Collection, actifs As Collection, a As Object
    On Error Resume Next
    modBaseIO.AssurerColonnes modConfig.FichierNomenclature(), "ACTES", EntetesNomenclature()
    On Error GoTo 0
    Set tous = modBaseIO.LireTableX(modConfig.FichierNomenclature(), "ACTES", "Code")
    Set actifs = New Collection
    For Each a In tous
        If a("Actif") <> "0" Then actifs.Add a
    Next a
    Set Nomenclature = actifs
End Function

Public Function ActeParCode(ByVal code As String) As Object
    Dim a As Object
    For Each a In Nomenclature()
        If a("Code") = code Then Set ActeParCode = a: Exit Function
    Next a
    Set ActeParCode = Nothing
End Function

' Code porte sur la feuille de soins : CodeOfficiel, sinon LibelleCerfa
' (ancienne colonne), sinon le code interne.
Public Function CodeOfficiel(ByVal a As Object) As String
    CodeOfficiel = Champ(a, "CodeOfficiel")
    If Len(CodeOfficiel) = 0 Then CodeOfficiel = Champ(a, "LibelleCerfa")
    If Len(CodeOfficiel) = 0 Then CodeOfficiel = Champ(a, "Code")
End Function

' ---------------------------------------------------------------------
' Controle de cotation : Config\regles_cotation.txt, une regle par ligne
'   INTERDIT=CODE+CODE;message        cumul interdit
'   UNIQUE=CODE;message               un seul exemplaire par seance
'   MAXPARSEANCE=n;message            nombre total d'actes par seance
' Renvoie "" si tout va bien, sinon la liste des avertissements.
' Ce controle ne remplace pas la verification de la cotation par le
' medecin : il signale les cas parametres par le cabinet.
' ---------------------------------------------------------------------
Public Function VerifierCotation(ByVal actesChoisis As Collection) As String
    Dim regles As Collection, r As Variant, parties() As String, corps As String
    Dim message As String, res As String, codes As Object, a As Object
    Dim membres() As String, i As Long, tousPresents As Boolean, n As Long

    Set codes = CreateObject("Scripting.Dictionary")
    codes.CompareMode = 1
    For Each a In actesChoisis
        codes(Champ(a, "Code")) = codes(Champ(a, "Code")) + 1
        n = n + 1
        If Len(Champ(a, "CodeAssocie")) > 0 Then
            codes(Champ(a, "CodeAssocie")) = codes(Champ(a, "CodeAssocie")) + 1
            n = n + 1
        End If
    Next a

    Set regles = LignesRegles()
    For Each r In regles
        parties = Split(CStr(r), "=")
        If UBound(parties) >= 1 Then
            corps = Mid$(CStr(r), Len(parties(0)) + 2)
            message = ""
            If InStr(corps, ";") > 0 Then
                message = Trim$(Mid$(corps, InStr(corps, ";") + 1))
                corps = Left$(corps, InStr(corps, ";") - 1)
            End If
            corps = Trim$(corps)
            Select Case UCase$(Trim$(parties(0)))
                Case "INTERDIT"
                    membres = Split(corps, "+")
                    tousPresents = True
                    For i = LBound(membres) To UBound(membres)
                        If Not codes.Exists(Trim$(membres(i))) Then tousPresents = False
                    Next i
                    If tousPresents And UBound(membres) >= 1 Then
                        res = res & "- cumul " & Replace(corps, "+", " + ") & " : " & _
                              IIf(Len(message) > 0, message, "association a verifier") & vbCrLf
                    End If
                Case "UNIQUE"
                    If codes.Exists(corps) Then
                        If codes(corps) > 1 Then
                            res = res & "- " & corps & " selectionne " & codes(corps) & " fois : " & _
                                  IIf(Len(message) > 0, message, "un seul par seance") & vbCrLf
                        End If
                    End If
                Case "MAXPARSEANCE"
                    If Val(corps) > 0 And n > Val(corps) Then
                        res = res & "- " & n & " actes pour cette seance (maximum parametre : " & _
                              corps & ") : " & IIf(Len(message) > 0, message, "a verifier") & vbCrLf
                    End If
            End Select
        End If
    Next r
    VerifierCotation = res
End Function

Private Function LignesRegles() As Collection
    Dim res As Collection, chemin As String, contenu As String, lignes() As String, i As Long, l As String
    Set res = New Collection
    chemin = modConfig.Chemin("Config") & "\regles_cotation.txt"
    Set LignesRegles = res
    If Not modFichiers.FichierExiste(chemin) Then Exit Function
    contenu = Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCrLf, vbLf), vbCr, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(lignes(i))
        If Len(l) > 0 And Left$(l, 1) <> "#" Then res.Add l
    Next i
End Function

' ---------------------------------------------------------------------
' Enregistrement d'une seance
' ---------------------------------------------------------------------
' infos       : dictionnaire du drapeau (SeanceID, ConsultationID, PatientID,
'               Nom, Prenom, DDN, NIR, DateActe)
' dateActe    : date REELLE de realisation (jj/mm/aaaa) ; jamais la date du
'               jour de traitement, qui peut etre le lendemain
' montantRegle: somme EFFECTIVEMENT recue ce jour (0 si rien n'est encaisse) ;
'               elle est repartie sur les lignes dans l'ordre
' Renvoie le SeanceID (stable : celui du drapeau s'il existe).
Public Function EnregistrerSeance(ByVal infos As Object, ByVal actesChoisis As Collection, _
                                  ByVal modePaiement As String, ByVal montantRegle As Double, _
                                  ByVal tiersPayant As Boolean, ByVal dateActe As String) As String
    Dim seanceID As String, lignes As Collection, a As Object, restant As Double
    seanceID = SeanceIDDe(infos)
    If Len(dateActe) = 0 Then dateActe = DateActeDe(infos)

    If modJournal.SeanceExiste(seanceID, AnneeDeDate(dateActe)) Then
        Err.Raise vbObjectError + 610, "modActes", _
            "Cette consultation est deja enregistree au journal (seance " & seanceID & ")." & vbCrLf & _
            "Utilisez 'Reimprimer la feuille' ou 'Enregistrer un reglement' : " & _
            "un nouveau document ne doit pas creer de nouveaux honoraires."
    End If

    restant = montantRegle
    Set lignes = New Collection
    For Each a In actesChoisis
        lignes.Add LigneJournal(seanceID, infos, dateActe, Champ(a, "Code"), CodeOfficiel(a), _
                                Champ(a, "Tarif"), modePaiement, tiersPayant, restant)
        If Len(Champ(a, "CodeAssocie")) > 0 Then
            lignes.Add LigneJournal(seanceID, infos, dateActe, Champ(a, "CodeAssocie"), Champ(a, "CodeAssocie"), _
                                    Champ(a, "TarifAssocie"), modePaiement, tiersPayant, restant)
        End If
    Next a
    modJournal.Ajouter lignes, AnneeDeDate(dateActe)
    EnregistrerSeance = seanceID
End Function

' Une ligne d'acte. restant est diminue du montant impute a cette ligne :
' un reglement partiel solde les premieres lignes, les suivantes restent dues.
Private Function LigneJournal(ByVal seanceID As String, ByVal infos As Object, _
                              ByVal dateActe As String, ByVal code As String, _
                              ByVal codeOff As String, ByVal tarif As String, _
                              ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                              ByRef restant As Double) As Object
    Dim d As Object, du As Double, regle As Double
    du = Val(Replace(tarif, ",", "."))
    If restant > 0 Then
        If restant >= du Then regle = du Else regle = restant
        restant = restant - regle
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d("DateActe") = dateActe
    d("DateSaisie") = Format$(Date, "dd/mm/yyyy")
    d("SeanceID") = seanceID
    d("ConsultationID") = ValeurOuVide(infos, "ConsultationID")
    d("PatientID") = ValeurOuVide(infos, "PatientID")
    d("Nom") = ValeurOuVide(infos, "Nom")
    d("Prenom") = ValeurOuVide(infos, "Prenom")
    d("DDN") = ValeurOuVide(infos, "DDN")
    d("NIR") = ValeurOuVide(infos, "NIR")
    d("CodeActe") = code
    d("CodeOfficiel") = codeOff
    d("MontantDu") = Format$(du, "0.00")
    d("MontantRegle") = Format$(regle, "0.00")
    d("ModePaiement") = modePaiement
    d("TiersPayant") = IIf(tiersPayant, "O", "N")
    ' Paye ne depend PAS du mode de paiement : seule une somme recue le met a O
    d("Paye") = IIf(regle >= du And du > 0, "O", "N")
    d("DateEncaissement") = IIf(regle > 0, Format$(Date, "dd/mm/yyyy"), "")
    d("FeuilleSoinsEtat") = ""
    d("DateFeuilleSoins") = ""
    Set LigneJournal = d
End Function

' Reglement recu apres coup : impute le montant sur les lignes non soldees
' de la seance, dans l'ordre. Renvoie le solde restant du apres imputation.
Public Function EnregistrerReglement(ByVal seanceID As String, ByVal montant As Double, _
                                     ByVal modePaiement As String, ByVal annee As Long) As Double
    Dim lignes As Collection, l As Object, restant As Double, solde As Double
    Dim du As Double, regle As Double, ajout As Double, v As Object
    Set lignes = modJournal.LignesSeance(seanceID, annee)
    If lignes.Count = 0 Then
        Err.Raise vbObjectError + 611, "modActes", "Seance introuvable au journal : " & seanceID
    End If
    restant = montant
    For Each l In lignes
        du = Val(Replace(ValeurOuVide(l, "MontantDu"), ",", "."))
        regle = Val(Replace(ValeurOuVide(l, "MontantRegle"), ",", "."))
        ajout = 0
        If restant > 0 And regle < du Then
            If restant >= (du - regle) Then ajout = du - regle Else ajout = restant
            restant = restant - ajout
        End If
        solde = solde + (du - regle - ajout)
    Next l
    ' le journal est tenu ligne par ligne, mais la mise a jour porte sur la
    ' seance : on reecrit le total regle et l'etat de paiement de la seance
    Set v = CreateObject("Scripting.Dictionary")
    v("ModePaiement") = modePaiement
    v("DateEncaissement") = Format$(Date, "dd/mm/yyyy")
    v("Paye") = IIf(solde <= 0.001, "O", "N")
    v("Notes") = "Reglement " & Format$(montant, "0.00") & " EUR le " & Format$(Date, "dd/mm/yyyy") & _
                 IIf(solde > 0.001, " - solde du " & Format$(solde, "0.00") & " EUR", "")
    modJournal.MettreAJourSeance seanceID, v, annee
    EnregistrerReglement = solde
End Function

Public Function SeanceIDDe(ByVal infos As Object) As String
    SeanceIDDe = ValeurOuVide(infos, "SeanceID")
    If Len(SeanceIDDe) = 0 Then SeanceIDDe = ValeurOuVide(infos, "ConsultationID")
    If Len(SeanceIDDe) = 0 Then SeanceIDDe = modFichiers.IdUnique()
End Function

Public Function DateActeDe(ByVal infos As Object) As String
    DateActeDe = ValeurOuVide(infos, "DateActe")
    If Len(DateActeDe) = 0 Then
        ' repli : date de validation du courrier par le medecin (jj/mm/aaaa hh:nn)
        DateActeDe = Left$(ValeurOuVide(infos, "DateValidation"), 10)
    End If
    If Len(DateActeDe) <> 10 Then DateActeDe = Format$(Date, "dd/mm/yyyy")
End Function

Public Function AnneeDeDate(ByVal dateTexte As String) As Long
    Dim parties() As String
    parties = Split(dateTexte, "/")
    If UBound(parties) = 2 Then AnneeDeDate = Val(parties(2)) Else AnneeDeDate = Year(Date)
End Function

Private Function ValeurOuVide(ByVal dict As Object, ByVal cle As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(cle) Then ValeurOuVide = CStr(dict(cle)) Else ValeurOuVide = ""
End Function

Private Function Champ(ByVal dict As Object, ByVal cle As String) As String
    Champ = ValeurOuVide(dict, cle)
End Function

' Lignes pour la feuille de soins : un element par acte, associes compris.
' dateActe : date REELLE de l'acte, portee sur chaque ligne du Cerfa.
Public Function LignesPourImpression(ByVal actesChoisis As Collection, ByVal dateActe As String) As Collection
    Dim col As New Collection, a As Object, d As Object
    For Each a In actesChoisis
        Set d = CreateObject("Scripting.Dictionary")
        d("CodeActe") = CodeOfficiel(a)
        d("Montant") = Champ(a, "Tarif")
        d("DateActe") = dateActe
        col.Add d
        If Len(Champ(a, "CodeAssocie")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d("CodeActe") = Champ(a, "CodeAssocie")
            d("Montant") = Champ(a, "TarifAssocie")
            d("DateActe") = dateActe
            col.Add d
        End If
    Next a
    Set LignesPourImpression = col
End Function

' Lignes de feuille de soins reconstituees depuis le journal (reimpression :
' memes actes, memes honoraires, meme date, aucun nouvel honoraire).
Public Function LignesDepuisJournal(ByVal seanceID As String, ByVal annee As Long) As Collection
    Dim col As New Collection, l As Object, d As Object
    For Each l In modJournal.LignesSeance(seanceID, annee)
        Set d = CreateObject("Scripting.Dictionary")
        d("CodeActe") = IIf(Len(ValeurOuVide(l, "CodeOfficiel")) > 0, ValeurOuVide(l, "CodeOfficiel"), ValeurOuVide(l, "CodeActe"))
        d("Montant") = ValeurOuVide(l, "MontantDu")
        d("DateActe") = ValeurOuVide(l, "DateActe")
        col.Add d
    Next l
    Set LignesDepuisJournal = col
End Function

' Total d'une selection d'actes (associes compris)
Public Function TotalActes(ByVal actesChoisis As Collection) As Double
    Dim a As Object, total As Double
    For Each a In actesChoisis
        total = total + Val(Replace(Champ(a, "Tarif"), ",", "."))
        If Len(Champ(a, "CodeAssocie")) > 0 Then total = total + Val(Replace(Champ(a, "TarifAssocie"), ",", "."))
    Next a
    TotalActes = total
End Function
