Option Explicit
' Traitement d'un document valide par le medecin.
'
' Un document n'est pas un honoraire : plusieurs documents (courrier
' principal + lettres de demande) proviennent de la MEME consultation.
' Le premier est enregistre avec ses actes ; les suivants sont traites
' par "Document seul (aucun acte)". Une seance deja enregistree ne peut
' pas l'etre deux fois : seuls le reglement et la reimpression restent
' accessibles.
Private mDrapeau As Object          ' dictionnaire du fichier-drapeau
Private mNomenclature As Collection
Private mInit As Boolean            ' vrai pendant le remplissage de la liste
Private mSeanceID As String
Private mDejaEnregistree As Boolean

Public Sub Charger(ByVal drapeau As Object)
    Dim modes As Variant, m As Variant, dateActe As String
    Set mDrapeau = drapeau
    mSeanceID = modActes.SeanceIDDe(drapeau)
    dateActe = modActes.DateActeDe(drapeau)
    mDejaEnregistree = SeanceDejaEnregistree(dateActe)

    lblInfo.Caption = Valeur("Prenom") & " " & Valeur("Nom") & _
                      IIf(Len(Valeur("DDN")) > 0, " - " & Valeur("DDN"), "") & vbCrLf & _
                      "Document : " & Valeur("TypeCourrier") & "  (valide le " & Valeur("DateValidation") & ")" & vbCrLf & _
                      IIf(mDejaEnregistree, _
                          ">>> Seance DEJA enregistree au journal (" & mSeanceID & ") : traitez ce document seul. <<<", _
                          "Seance " & mSeanceID)
    lstItemsInit
    cmbPaiement.Clear
    modes = Array("CB", "Cheque", "Especes", "Virement attendu", "Impaye")
    For Each m In modes
        cmbPaiement.AddItem CStr(m)
    Next m
    cmbPaiement.ListIndex = 0
    chkFds.Value = Not mDejaEnregistree
    txtDateActe.Text = dateActe
    btnOK.Enabled = Not mDejaEnregistree
    MajTotal
End Sub

' Une seance est deja au journal si ses lignes existent (reprise apres
' erreur d'impression, deuxieme document de la meme consultation...).
Private Function SeanceDejaEnregistree(ByVal dateActe As String) As Boolean
    On Error Resume Next
    SeanceDejaEnregistree = modJournal.SeanceExiste(mSeanceID, modActes.AnneeDeDate(dateActe))
    On Error GoTo 0
End Function

Private Sub lstItemsInit()
    Dim a As Object, i As Long, typeCourrier As String
    mInit = True
    Set mNomenclature = modActes.Nomenclature()
    lstActes.Clear
    lstActes.ColumnCount = 3
    lstActes.ColumnWidths = "60 pt;250 pt;60 pt"
    lstActes.MultiSelect = 1            ' fmMultiSelectMulti
    lstActes.ListStyle = 1              ' fmListStyleOption (cases a cocher)
    i = 0
    typeCourrier = LCase$(Valeur("TypeCourrier"))
    For Each a In mNomenclature
        lstActes.AddItem ""
        lstActes.List(i, 0) = a("Code")
        lstActes.List(i, 1) = a("LibelleCourt") & IIf(Len(a("CodeAssocie")) > 0, " (+ " & a("CodeAssocie") & ")", "")
        lstActes.List(i, 2) = Format$(Val(Replace(a("Tarif"), ",", ".")) + Val(Replace(a("TarifAssocie"), ",", ".")), "0.00")
        ' preselection simple : consultation -> CSC ; jamais pour une lettre de demande
        If Not mDejaEnregistree And InStr(typeCourrier, "consultation") > 0 And a("Code") = "CSC" Then
            lstActes.Selected(i) = True
        End If
        i = i + 1
    Next a
    mInit = False
End Sub

Private Function Valeur(ByVal cle As String) As String
    If mDrapeau.Exists(cle) Then Valeur = mDrapeau(cle) Else Valeur = ""
End Function

Private Function ActesChoisis() As Collection
    Dim col As New Collection, j As Long, a As Object
    If mNomenclature Is Nothing Then Set ActesChoisis = col: Exit Function
    j = 0
    For Each a In mNomenclature
        If j < lstActes.ListCount Then
            If lstActes.Selected(j) Then col.Add a
        End If
        j = j + 1
    Next a
    Set ActesChoisis = col
End Function

Private Sub MajTotal()
    On Error Resume Next
    Dim total As Double
    total = modActes.TotalActes(ActesChoisis())
    lblTotal.Caption = "Total du : " & Format$(total, "0.00") & " EUR"
    ' la somme recue suit le total tant que la secretaire ne l'a pas modifiee
    If Not mInit And EstEncaissementImmediat() Then txtRegle.Text = Format$(total, "0.00")
End Sub

Private Function EstEncaissementImmediat() As Boolean
    Dim m As String
    m = LCase$(cmbPaiement.Text)
    EstEncaissementImmediat = (m = "cb" Or m = "cheque" Or m = "especes")
End Function

Private Sub lstActes_Change()
    If mInit Then Exit Sub
    MajTotal
End Sub

Private Sub cmbPaiement_Change()
    If mInit Then Exit Sub
    ' un virement attendu ou un impaye n'encaisse rien tant qu'il n'est pas recu
    If EstEncaissementImmediat() Then
        txtRegle.Text = Format$(modActes.TotalActes(ActesChoisis()), "0.00")
    Else
        txtRegle.Text = "0.00"
    End If
End Sub

Private Sub btnOuvrir_Click()
    modEchange.OuvrirCourrier mDrapeau
End Sub

' Traite le document sans creer d'honoraire : lettre de demande, second
' courrier de la meme consultation, document a classer seulement.
Private Sub btnDocSeul_Click()
    On Error GoTo Erreur
    If MsgBox("Classer ce document sans enregistrer d'acte ni d'honoraire ?" & vbCrLf & vbCrLf & _
              Valeur("TypeCourrier") & " - " & Valeur("Prenom") & " " & Valeur("Nom"), _
              vbYesNo + vbQuestion, "Cabinet") <> vbYes Then Exit Sub
    If mDrapeau.Exists("_Chemin") Then modEchange.DeplacerVersTraites mDrapeau("_Chemin")
    MsgBox "Document classe. Aucun acte, aucun honoraire enregistre.", vbInformation, "Cabinet"
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim actes As Collection, a As Object, tarifZero As Boolean
    Dim dateActe As String, montantRegle As Double, total As Double, avert As String

    If mDejaEnregistree Then
        MsgBox "Cette seance est deja au journal." & vbCrLf & _
               "Utilisez 'Document seul', 'Enregistrer un reglement' ou 'Reimprimer la feuille'.", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If

    Set actes = ActesChoisis()
    If actes.Count = 0 Then
        MsgBox "Cochez au moins un acte, ou utilisez 'Document seul (aucun acte)' " & _
               "pour un document qui ne cree pas d'honoraire.", vbExclamation, "Cabinet"
        Exit Sub
    End If

    dateActe = Trim$(txtDateActe.Text)
    If Not modTexte.DateFrValide(dateActe) Then
        MsgBox "Date de realisation invalide (jj/mm/aaaa).", vbExclamation, "Cabinet"
        Exit Sub
    End If

    For Each a In actes
        If Val(Replace(a("Tarif"), ",", ".")) = 0 Then tarifZero = True
    Next a
    If tarifZero Then
        If MsgBox("Au moins un acte a un tarif a 0 EUR (nomenclature a completer : " & _
                  "Config\Nomenclature.xlsx)." & vbCrLf & "Continuer quand meme ?", _
                  vbYesNo + vbExclamation, "Cabinet") <> vbYes Then Exit Sub
    End If

    ' controle de cotation parametre par le cabinet (Config\regles_cotation.txt)
    avert = modActes.VerifierCotation(actes)
    If Len(avert) > 0 Then
        If MsgBox("Cotation a verifier :" & vbCrLf & vbCrLf & avert & vbCrLf & _
                  "Enregistrer quand meme ?", vbYesNo + vbExclamation, "Cabinet") <> vbYes Then Exit Sub
    End If

    total = modActes.TotalActes(actes)
    montantRegle = Val(Replace(Trim$(txtRegle.Text), ",", "."))
    If montantRegle < 0 Then montantRegle = 0
    If montantRegle > total + 0.001 Then
        MsgBox "La somme recue (" & Format$(montantRegle, "0.00") & " EUR) depasse le total du (" & _
               Format$(total, "0.00") & " EUR).", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If chkTiers.Value And montantRegle > 0 Then
        If MsgBox("Tiers payant coche ET " & Format$(montantRegle, "0.00") & " EUR recus du patient." & vbCrLf & _
                  "Est-ce bien la part effectivement payee ?", vbYesNo + vbQuestion, "Cabinet") <> vbYes Then Exit Sub
    End If

    ' 1. journal : honoraires dus et somme reellement recue (feuille de soins
    '    a l'etat "aucune" tant qu'elle n'a pas ete envoyee a l'imprimante)
    mSeanceID = modActes.EnregistrerSeance(mDrapeau, actes, cmbPaiement.Text, montantRegle, _
                                           chkTiers.Value, dateActe)
    mDejaEnregistree = True
    btnOK.Enabled = False

    ' 2. feuille de soins : demandee -> envoyee -> sortie confirmee par la secretaire
    If chkFds.Value Then ImprimerEtConfirmer actes, dateActe

    ' 3. le document quitte la file du secretariat
    If mDrapeau.Exists("_Chemin") Then modEchange.DeplacerVersTraites mDrapeau("_Chemin")

    MsgBox "Seance " & mSeanceID & " enregistree." & vbCrLf & _
           "Du : " & Format$(total, "0.00") & " EUR   Recu : " & Format$(montantRegle, "0.00") & " EUR" & _
           IIf(total - montantRegle > 0.001, "   Reste du : " & Format$(total - montantRegle, "0.00") & " EUR", ""), _
           vbInformation, "Cabinet"
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Imprime la feuille et enregistre son etat REEL. L'absence d'erreur Windows
' ne prouve pas la sortie papier : la secretaire confirme (ou non).
Private Sub ImprimerEtConfirmer(ByVal actes As Collection, ByVal dateActe As String)
    Dim etat As String, v As Object, lignes As Collection
    MarquerFeuille "Demandee", dateActe
    On Error GoTo EchecImpression
    Set lignes = modActes.LignesPourImpression(actes, dateActe)
    modCerfaPrint.ImprimerFeuille InfosCerfa(dateActe), lignes
    On Error GoTo 0
    If MsgBox("La feuille de soins est-elle sortie correctement et complete ?" & vbCrLf & vbCrLf & _
              "Non : elle restera 'a reimprimer' (aucun honoraire n'est modifie).", _
              vbYesNo + vbQuestion, "Feuille de soins") = vbYes Then
        etat = "Remise"
    Else
        etat = "EchecImpression"
    End If
    MarquerFeuille etat, dateActe
    Exit Sub
EchecImpression:
    Dim descErr As String
    descErr = Err.Description
    On Error Resume Next
    MarquerFeuille "EchecImpression", dateActe
    On Error GoTo 0
    MsgBox "Impression impossible : " & descErr & vbCrLf & vbCrLf & _
           "La seance reste enregistree ; utilisez 'Reimprimer la feuille' apres correction.", _
           vbExclamation, "Cabinet"
End Sub

Private Sub MarquerFeuille(ByVal etat As String, ByVal dateActe As String)
    Dim v As Object
    Set v = CreateObject("Scripting.Dictionary")
    v("FeuilleSoinsEtat") = etat
    v("DateFeuilleSoins") = Format$(Now, "dd/mm/yyyy hh:nn")
    On Error Resume Next
    modJournal.MettreAJourSeance mSeanceID, v, modActes.AnneeDeDate(dateActe)
    On Error GoTo 0
End Sub

' Informations d'identite pour le Cerfa : patient, assure s'il differe,
' medecin traitant pour la rubrique parcours de soins.
Private Function InfosCerfa(ByVal dateActe As String) As Object
    Dim d As Object, k As Variant, pat As Object, cor As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For Each k In mDrapeau.Keys
        d(CStr(k)) = mDrapeau(k)
    Next k
    d("DateActe") = dateActe
    On Error Resume Next
    For Each pat In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If pat("ID") = Valeur("PatientID") Then
            ' assure distinct s'il est renseigne dans la fiche patient
            If pat.Exists("AssureNom") Then
                If Len(pat("AssureNom")) > 0 Then
                    d("AssureNom") = pat("AssureNom")
                    d("AssurePrenom") = pat("AssurePrenom")
                    d("AssureDDN") = pat("AssureDDN")
                    d("AssureNIR") = pat("AssureNIR")
                End If
            End If
            If Len(pat("MedTraitantID")) > 0 Then
                For Each cor In modBaseIO.LireTableX(modConfig.FichierPatients(), "CORRESPONDANTS")
                    If cor("ID") = pat("MedTraitantID") Then
                        d("MedTraitantNom") = Trim$(cor("Nom") & " " & cor("Prenom"))
                        Exit For
                    End If
                Next cor
            End If
            Exit For
        End If
    Next pat
    On Error GoTo 0
    Set InfosCerfa = d
End Function

' Reglement recu apres coup : n'ajoute aucun acte, aucun honoraire.
Private Sub btnReglement_Click()
    On Error GoTo Erreur
    Dim rep As String, montant As Double, solde As Double, dateActe As String
    dateActe = Trim$(txtDateActe.Text)
    If Not mDejaEnregistree Then
        MsgBox "Enregistrez d'abord la seance.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    rep = InputBox("Somme recue (EUR) pour la seance " & mSeanceID & " :", "Reglement", "0.00")
    If Len(rep) = 0 Then Exit Sub
    montant = Val(Replace(rep, ",", "."))
    If montant <= 0 Then Exit Sub
    solde = modActes.EnregistrerReglement(mSeanceID, montant, cmbPaiement.Text, modActes.AnneeDeDate(dateActe))
    MsgBox "Reglement de " & Format$(montant, "0.00") & " EUR enregistre." & vbCrLf & _
           IIf(solde > 0.001, "Reste du : " & Format$(solde, "0.00") & " EUR.", "Seance soldee."), _
           vbInformation, "Cabinet"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Reimpression : memes actes, memes honoraires, meme date, relus au
' journal. Aucune ligne financiere n'est ajoutee.
Private Sub btnReimprimer_Click()
    On Error GoTo Erreur
    Dim dateActe As String, lignes As Collection
    dateActe = Trim$(txtDateActe.Text)
    If Not mDejaEnregistree Then
        MsgBox "Aucune seance enregistree pour ce document.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set lignes = modActes.LignesDepuisJournal(mSeanceID, modActes.AnneeDeDate(dateActe))
    If lignes.Count = 0 Then
        MsgBox "Aucune ligne d'acte trouvee au journal pour la seance " & mSeanceID & ".", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    MarquerFeuille "Demandee", dateActe
    modCerfaPrint.ImprimerFeuille InfosCerfa(dateActe), lignes
    If MsgBox("La feuille est-elle sortie correctement ?", vbYesNo + vbQuestion, "Feuille de soins") = vbYes Then
        MarquerFeuille "Remise", dateActe
        MsgBox "Reimpression enregistree (aucun nouvel honoraire).", vbInformation, "Cabinet"
    Else
        MarquerFeuille "EchecImpression", dateActe
    End If
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    On Error Resume Next
    MarquerFeuille "EchecImpression", Trim$(txtDateActe.Text)
    On Error GoTo 0
    MsgBox "Erreur : " & descErr, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
