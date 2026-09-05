Option Explicit
' Prise de rendez-vous (poste secretaire).
Private mPatient As Object

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
    If mPatient Is Nothing Then
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
    conflit = modAgenda.ConflitRdv(Trim$(txtDate.Text), Trim$(txtHeure.Text), cmbDuree.Text)
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
    id = modAgenda.AjouterRdv(mPatient("ID"), Trim$(txtDate.Text), Trim$(txtHeure.Text), _
                              cmbDuree.Text, cmbType.Text, Trim$(txtNotes.Text), forcer)
    MsgBox "RDV enregistre (" & id & ") : " & mPatient("Prenom") & " " & mPatient("Nom") & _
           " le " & txtDate.Text & " a " & txtHeure.Text & ".", vbInformation, "Cabinet"
    Me.Hide
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub btnAnnuler_Click()
    Me.Hide
End Sub
