Option Explicit
' Fiche patient : creation / modification (poste secretaire).
Private mID As String               ' vide = nouveau patient
Private mIDsMedecins() As String

Public Sub ChargerNouveau()
    mID = ""
    Me.Caption = "Nouveau patient"
    lblID.Caption = ""
    ChargerCombos
    cmbSexe.ListIndex = 0
End Sub

Public Sub ChargerExistant(ByVal pat As Object)
    mID = pat("ID")
    Me.Caption = "Patient " & pat("Prenom") & " " & pat("Nom")
    lblID.Caption = "Dossier " & mID
    ChargerCombos
    txtNom.Text = pat("Nom")
    txtNomN.Text = pat("NomNaissance")
    txtPrenom.Text = pat("Prenom")
    txtDDN.Text = modTexte.DdnPatient(pat)
    SelectionnerCombo cmbSexe, modTexte.SexePatient(pat)
    txtNIR.Text = pat("NIR")
    txtAdr1.Text = pat("Adresse1")
    txtAdr2.Text = pat("Adresse2")
    txtCP.Text = pat("CP")
    txtVille.Text = pat("Ville")
    txtTel.Text = pat("Tel")
    txtMobile.Text = pat("Mobile")
    txtEmail.Text = pat("Email")
    SelectionnerMedecin pat("MedTraitantID")
    txtMutuelle.Text = pat("Mutuelle")
    chkALD.Value = (pat("ALD") = "O")
    txtNotes.Text = pat("Notes")
End Sub

Private Sub ChargerCombos()
    Dim cors As Collection, c As Object, i As Long
    cmbSexe.Clear
    cmbSexe.AddItem "M"
    cmbSexe.AddItem "F"
    ' Medecins traitants : tries par NOM, libelle "NOM Prenom - Ville" pour
    ' que les premieres lettres tapees dans la liste tombent sur le nom
    ' (MatchEntry = complete : la saisie se complete au fur et a mesure).
    cmbMed.Clear
    cmbMed.AddItem "(aucun)"
    Set cors = TrierParNom(modBaseIO.LireTableX(modConfig.FichierPatients(), "CORRESPONDANTS"))
    ReDim mIDsMedecins(0 To cors.Count)
    mIDsMedecins(0) = ""
    i = 1
    For Each c In cors
        cmbMed.AddItem UCase$(Trim$(c("Nom"))) & " " & Trim$(c("Prenom")) & IIf(Len(Trim$(c("Ville"))) > 0, " - " & Trim$(c("Ville")), "")
        mIDsMedecins(i) = c("ID")
        i = i + 1
    Next c
    cmbMed.ListIndex = 0
End Sub

' Correspondants actifs, tries par Nom puis Prenom (insertion triee)
Private Function TrierParNom(ByVal cors As Collection) As Collection
    Dim res As Collection, c As Object, i As Long, cle As String, inserted As Boolean
    Set res = New Collection
    For Each c In cors
        If c("Actif") <> "0" Then
            cle = UCase$(Trim$(c("Nom")) & " " & Trim$(c("Prenom")))
            inserted = False
            For i = 1 To res.Count
                If UCase$(Trim$(res(i)("Nom")) & " " & Trim$(res(i)("Prenom"))) > cle Then
                    res.Add c, , i: inserted = True: Exit For
                End If
            Next i
            If Not inserted Then res.Add c
        End If
    Next c
    Set TrierParNom = res
End Function

' Texte tape sans selection dans la liste : on retrouve l'entree qui commence
' par ce texte (premieres lettres du nom) ; sinon on refuse d'enregistrer.
Private Function IndexMedecinSaisi() As Long
    Dim i As Long, t As String
    IndexMedecinSaisi = cmbMed.ListIndex
    If IndexMedecinSaisi >= 0 Then Exit Function
    t = UCase$(Trim$(cmbMed.Text))
    If Len(t) = 0 Then IndexMedecinSaisi = 0: Exit Function
    For i = 1 To cmbMed.ListCount - 1
        If Left$(UCase$(cmbMed.List(i)), Len(t)) = t Then IndexMedecinSaisi = i: Exit Function
    Next i
    IndexMedecinSaisi = -1
End Function

Private Sub SelectionnerMedecin(ByVal id As String)
    Dim i As Long
    For i = 0 To UBound(mIDsMedecins)
        If mIDsMedecins(i) = id Then
            cmbMed.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub SelectionnerCombo(ByVal cmb As MSForms.ComboBox, ByVal valeur As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = valeur Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim d As Object, nouveauID As String
    If Len(Trim$(txtNom.Text)) = 0 Or Len(Trim$(txtPrenom.Text)) = 0 Then
        MsgBox "Le nom et le prenom sont obligatoires.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Not DateValide(txtDDN.Text) Then
        MsgBox "Date de naissance invalide : jj/mm/aaaa, date reelle du calendrier, non posterieure a aujourd'hui.", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d("Nom") = UCase$(Trim$(txtNom.Text))
    d("NomNaissance") = UCase$(Trim$(txtNomN.Text))
    d("Prenom") = Trim$(txtPrenom.Text)
    d("DDN") = Trim$(txtDDN.Text)
    d("Sexe") = cmbSexe.Text
    d("NIR") = Trim$(txtNIR.Text)
    d("Adresse1") = Trim$(txtAdr1.Text)
    d("Adresse2") = Trim$(txtAdr2.Text)
    d("CP") = Trim$(txtCP.Text)
    d("Ville") = Trim$(txtVille.Text)
    d("Tel") = Trim$(txtTel.Text)
    d("Mobile") = Trim$(txtMobile.Text)
    d("Email") = Trim$(txtEmail.Text)
    Dim iMed As Long
    iMed = IndexMedecinSaisi()
    If iMed < 0 Then
        MsgBox "Medecin traitant introuvable : '" & cmbMed.Text & "'. Tapez les premieres lettres du NOM et choisissez dans la liste.", _
               vbExclamation, "Cabinet"
        cmbMed.SetFocus
        Exit Sub
    End If
    d("MedTraitantID") = mIDsMedecins(iMed)
    d("Mutuelle") = Trim$(txtMutuelle.Text)
    d("ALD") = IIf(chkALD.Value, "O", "")
    d("Notes") = Trim$(txtNotes.Text)
    If Len(mID) = 0 Then
        d("DateCreation") = Format$(Now, "dd/mm/yyyy")
        nouveauID = modBaseIO.AjouterLigne(modConfig.FichierPatients(), "PATIENTS", d, "P")
        MsgBox "Patient cree : " & d("Prenom") & " " & d("Nom") & " (" & nouveauID & ")", _
               vbInformation, "Cabinet"
    Else
        d("DateModif") = Format$(Now, "dd/mm/yyyy")
        ' fiche provisoire (creee depuis l'agenda) : complete des que la
        ' date de naissance est renseignee
        If Len(Trim$(txtDDN.Text)) > 0 Then d("Provisoire") = ""
        modBaseIO.ModifierLigne modConfig.FichierPatients(), "PATIENTS", "ID", mID, d
        MsgBox "Fiche mise a jour.", vbInformation, "Cabinet"
    End If
    On Error Resume Next
    modAgendaVue.InvaliderPatients
    On Error GoTo 0
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Validation calendaire stricte (modTexte) : le 31 fevrier est refuse,
' et une date de naissance posterieure a aujourd'hui aussi.
Private Function DateValide(ByVal s As String) As Boolean
    If Not modTexte.DateFrValide(s) Then Exit Function
    If modTexte.DateFr(s) > Date Then Exit Function
    DateValide = True
End Function

Private Sub btnAgenda_Click()
    Dim d As Object
    If Len(mID) = 0 Then
        MsgBox "Enregistrez d'abord la fiche du patient, puis reprenez-la pour fixer le rendez-vous.", _
               vbInformation, "Cabinet"
        Exit Sub
    End If
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = mID
    d("Nom") = UCase$(Trim$(txtNom.Text))
    d("Prenom") = Trim$(txtPrenom.Text)
    d("DDN") = Trim$(txtDDN.Text)
    Me.Hide
    modAgendaVue.AfficherAgenda Date, d
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
