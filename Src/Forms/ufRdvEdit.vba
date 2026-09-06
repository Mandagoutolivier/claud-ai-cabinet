Option Explicit
' Prise de rendez-vous (poste secretaire). Sert aussi a DEPLACER / modifier
' un rendez-vous existant (ChargerExistant) : meme ecran, meme controles.
Private mPatient As Object
Private mRdvID As String          ' vide = creation ; sinon modification
Private mAnneeOrigine As Long

Private Sub UserForm_Initialize()
    Dim a As Object, duree As Variant
    txtDate.Text = Format$(Date, "dd/mm/yyyy")
    txtHeure.Text = ""
    ' valeurs par defaut du cabinet : config.ini [AGENDA] DureeDefautMin / ActeDefaut
    Dim dureeDef As String, acteDef As String, i As Long
    dureeDef = CStr(modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15))
    acteDef = modConfig.Config("AGENDA", "ActeDefaut", "APC")
    For Each duree In Array("15", "20", "30", "45", "60")
        cmbDuree.AddItem CStr(duree)
    Next duree
    cmbDuree.ListIndex = 0
    For i = 0 To cmbDuree.ListCount - 1
        If cmbDuree.List(i) = dureeDef Then cmbDuree.ListIndex = i
    Next i
    cmbType.Clear
    On Error Resume Next
    For Each a In modActes.Nomenclature()
        cmbType.AddItem a("Code")
    Next a
    On Error GoTo 0
    If cmbType.ListCount > 0 Then cmbType.ListIndex = 0
    For i = 0 To cmbType.ListCount - 1
        If UCase$(cmbType.List(i)) = UCase$(acteDef) Then cmbType.ListIndex = i
    Next i
End Sub

' [AGENDA] DureeParActe=CS:15;CSC:20;ETT:30... : la duree suit le motif choisi
Private Sub cmbType_Change()
    Dim regle As Variant, parts() As String, i As Long
    On Error Resume Next
    For Each regle In Split(modConfig.Config("AGENDA", "DureeParActe", ""), ";")
        parts = Split(Trim$(regle), ":")
        If UBound(parts) = 1 Then
            If UCase$(Trim$(parts(0))) = UCase$(Trim$(cmbType.Text)) Then
                For i = 0 To cmbDuree.ListCount - 1
                    If cmbDuree.List(i) = Trim$(parts(1)) Then cmbDuree.ListIndex = i: Exit Sub
                Next i
                cmbDuree.AddItem Trim$(parts(1))
                cmbDuree.ListIndex = cmbDuree.ListCount - 1
                Exit Sub
            End If
        End If
    Next regle
End Sub

' Modification / deplacement d'un RDV existant (depuis l'agenda)
Public Sub ChargerExistant(ByVal rdv As Object, ByVal pat As Object)
    Dim i As Long
    mRdvID = rdv("ID")
    mAnneeOrigine = Val(Right$(rdv("Date"), 4))
    Me.Caption = "Deplacer / modifier le rendez-vous"
    btnOK.Caption = "Enregistrer la modification"
    btnPatient.Enabled = False
    txtDate.Text = rdv("Date")
    txtHeure.Text = rdv("Heure")
    txtNotes.Text = rdv("Notes")
    For i = 0 To cmbDuree.ListCount - 1
        If cmbDuree.List(i) = rdv("DureeMin") Then cmbDuree.ListIndex = i
    Next i
    If cmbDuree.Text <> rdv("DureeMin") And Val(rdv("DureeMin")) > 0 Then
        cmbDuree.AddItem rdv("DureeMin"): cmbDuree.ListIndex = cmbDuree.ListCount - 1
    End If
    For i = 0 To cmbType.ListCount - 1
        If UCase$(cmbType.List(i)) = UCase$(rdv("TypeActe")) Then cmbType.ListIndex = i
    Next i
    If Not pat Is Nothing Then
        Set mPatient = pat
        lblPatient.Caption = pat("Prenom") & " " & pat("Nom") & " (" & pat("DDN") & ")"
    Else
        lblPatient.Caption = "Patient " & rdv("PatientID")
    End If
End Sub

' Pre-remplissage depuis l'agenda (double-clic sur un creneau libre)
Public Sub Prefixer(ByVal dateTxt As String, ByVal heureTxt As String, Optional ByVal pat As Object = Nothing)
    txtDate.Text = dateTxt
    txtHeure.Text = heureTxt
    If Not pat Is Nothing Then
        Set mPatient = pat
        lblPatient.Caption = pat("Prenom") & " " & pat("Nom") & " (" & pat("DDN") & ")"
    End If
End Sub

Private Sub btnAgenda_Click()
    Dim d As Date
    On Error Resume Next
    d = CDate(Trim$(txtDate.Text))
    If Err.Number <> 0 Then d = Date
    On Error GoTo 0
    Me.Hide
    modAgendaVue.AfficherAgenda d, mPatient
End Sub

Private Sub btnPatient_Click()
    Dim p As Object
    Set p = modUI.ChoisirPatientX()
    If Not p Is Nothing Then
        Set mPatient = p
        lblPatient.Caption = p("Prenom") & " " & p("Nom") & " (" & p("DDN") & ")"
    End If
End Sub

Private Sub btnOK_Click()
    On Error GoTo Erreur
    Dim id As String, conflit As String, forcer As Boolean
    If mPatient Is Nothing And Len(mRdvID) = 0 Then
        MsgBox "Choisissez d'abord le patient.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    If Not modTexte.DateFrValide(Trim$(txtDate.Text)) Then
        MsgBox "Date invalide : saisissez une date existante au format jj/mm/aaaa.", vbExclamation, "Cabinet"
        txtDate.SetFocus
        Exit Sub
    End If
    If Not modTexte.HeureValide(Trim$(txtHeure.Text)) Then
        MsgBox "Heure invalide : saisissez hh:mm entre 00:00 et 23:59.", vbExclamation, "Cabinet"
        txtHeure.SetFocus
        Exit Sub
    End If
    If Val(cmbDuree.Text) <= 0 Then
        MsgBox "Duree invalide : choisissez une duree en minutes.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    conflit = modAgenda.ConflitRdv(Trim$(txtDate.Text), Trim$(txtHeure.Text), cmbDuree.Text, mRdvID)
    If Len(conflit) > 0 Then
        If MsgBox("Ce creneau chevauche un rendez-vous deja prevu le " & txtDate.Text & _
                  " a " & conflit & "." & vbCrLf & vbCrLf & _
                  "Enregistrer quand meme (double creneau) ?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Cabinet") <> vbYes Then
            txtHeure.SetFocus
            Exit Sub
        End If
        forcer = True
    End If
    If Len(mRdvID) > 0 Then
        id = modAgenda.ModifierRdv(mRdvID, mAnneeOrigine, Trim$(txtDate.Text), Trim$(txtHeure.Text), _
                                   cmbDuree.Text, cmbType.Text, Trim$(txtNotes.Text), forcer)
        MsgBox "RDV modifie : " & lblPatient.Caption & " le " & txtDate.Text & " a " & txtHeure.Text & ".", _
               vbInformation, "Cabinet"
    Else
        id = modAgenda.AjouterRdv(mPatient("ID"), Trim$(txtDate.Text), Trim$(txtHeure.Text), _
                                  cmbDuree.Text, cmbType.Text, Trim$(txtNotes.Text), forcer)
        MsgBox "RDV enregistre (" & id & ") : " & mPatient("Prenom") & " " & mPatient("Nom") & _
               " le " & txtDate.Text & " a " & txtHeure.Text & ".", vbInformation, "Cabinet"
    End If
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
