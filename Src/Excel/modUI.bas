Attribute VB_Name = "modUI"
Option Explicit
' =====================================================================
' modUI - Actions des boutons de la feuille Accueil (poste secretaire).
' =====================================================================

Public Sub UI_NouveauPatient()
    On Error GoTo Erreur
    Dim f As ufPatientEdit
    Set f = New ufPatientEdit
    f.ChargerNouveau
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_ModifierPatient()
    On Error GoTo Erreur
    Dim pat As Object, f As ufPatientEdit
    Set pat = ChoisirPatientX()
    If pat Is Nothing Then Exit Sub
    Set f = New ufPatientEdit
    f.ChargerExistant pat
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_NouveauRdv()
    On Error GoTo Erreur
    Dim f As ufRdvEdit
    Set f = New ufRdvEdit
    f.Show vbModal
    Unload f
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_ArriveePatient()
    On Error GoTo Erreur
    Dim rdvs As Collection, f As ufListe, r As Object
    Set rdvs = modAgenda.RdvDuJour()
    If rdvs.Count = 0 Then
        MsgBox "Aucun rendez-vous aujourd'hui." & vbCrLf & _
               "Utilisez 'Prise de rendez-vous' pour en creer un.", vbInformation, "Cabinet"
        Exit Sub
    End If
    Set f = New ufListe
    f.Configurer "Arrivee d'un patient - RDV du jour", rdvs, _
                 Array("Heure", "Nom", "Prenom", "TypeActe", "Statut"), "40 pt;110 pt;90 pt;70 pt;60 pt"
    f.Show vbModal
    If f.Annule Then Unload f: Exit Sub
    Set r = f.Resultat
    Unload f
    modAgenda.MarquerStatut r("ID"), "Arrive"
    ' file des patients arrives pour le poste medecin (Ctrl+Alt+N)
    On Error Resume Next
    Dim pArr As Object
    For Each pArr In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If pArr("ID") = r("PatientID") Then modAgenda.SignalerArrivee pArr, r("ID"), r("Heure"): Exit For
    Next pArr
    If Err.Number <> 0 Then modLog.LogErreur "Signal d'arrivee : " & Err.Description
    On Error GoTo Erreur
    ' envoi automatique de l'identite au poste ECG (si configure)
    Dim noteEcg As String
    If Len(modConfig.Config("ECG", "DossierGdt", "")) > 0 Then
        On Error Resume Next
        Dim p As Object
        For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
            If p("ID") = r("PatientID") Then
                modGdt.EcrireGdtPatient p
                If Err.Number = 0 Then
                    noteEcg = vbCrLf & "Identite envoyee a l'ECG."
                Else
                    noteEcg = vbCrLf & "ATTENTION : identite NON envoyee a l'ECG - " & Err.Description & vbCrLf & _
                              "(verifier [ECG] DossierGdt dans Config\config.ini : chemin reseau vers le PC de l'ECG)"
                    modLog.LogErreur "Envoi ECG (arrivee) : " & Err.Description
                    Err.Clear
                End If
                Exit For
            End If
        Next p
        On Error GoTo Erreur
    End If
    MsgBox r("Prenom") & " " & r("Nom") & " marque ARRIVE a " & Format$(Now, "hh:nn") & "." & noteEcg, _
           vbInformation, "Cabinet"
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_Correspondants()
    On Error GoTo Erreur
    Dim f As ufListe, fc As ufCorrespEdit
    Set f = New ufListe
    f.Configurer "Correspondants (Nouveau... pour en creer un)", _
                 modBaseIO.LireTableX(modConfig.FichierPatients(), "CORRESPONDANTS"), _
                 Array("Titre", "Nom", "Prenom", "Specialite", "Ville"), _
                 "30 pt;110 pt;80 pt;120 pt;80 pt", "", True
    f.Show vbModal
    If f.NouveauDemande Then
        Unload f
        Set fc = New ufCorrespEdit
        fc.ChargerNouveau
        fc.Show vbModal
        Unload fc
    ElseIf Not f.Annule Then
        Dim cor As Object
        Set cor = f.Resultat
        Unload f
        Set fc = New ufCorrespEdit
        fc.ChargerExistant cor
        fc.Show vbModal
        Unload fc
    Else
        Unload f
    End If
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_CourriersATraiter()
    On Error GoTo Erreur
    Dim courriers As Collection, f As ufListe, d As Object, fa As ufChoixActe
    Set courriers = modEchange.CourriersEnAttente()
    If courriers.Count = 0 Then
        MsgBox "Aucun courrier en attente.", vbInformation, "Cabinet"
        Exit Sub
    End If
    Set f = New ufListe
    f.Configurer "Courriers valides par le medecin", courriers, _
                 Array("Nom", "Prenom", "TypeCourrier", "DateValidation"), "110 pt;90 pt;110 pt;90 pt"
    f.Show vbModal
    If f.Annule Then Unload f: Exit Sub
    Set d = f.Resultat
    Unload f
    Set fa = New ufChoixActe
    fa.Charger d
    fa.Show vbModal
    Unload fa
    modEchange.VerifierEchange
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_Agenda()
    On Error GoTo Erreur
    modAgendaVue.AfficherAgenda Date
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_Honoraires()
    On Error GoTo Erreur
    modHonoraires.Afficher
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Fiches patients incompletes pour l'ECG (date de naissance / sexe) : liste
' dans une feuille "Controle" pour completer au fil de l'eau.
Public Sub UI_ControleFiches()
    On Error GoTo Erreur
    Dim p As Object, ws As Worksheet, r As Long, nDdn As Long, nSexe As Long, valeurs As Object, k As Variant
    Set valeurs = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Controle")
    On Error GoTo Erreur
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Controle"
    End If
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Fiches incompletes pour l'ECG (date de naissance ou sexe) - " & Format$(Now, "dd/mm/yyyy hh:nn")
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 1).Value = "ID": ws.Cells(3, 2).Value = "Nom": ws.Cells(3, 3).Value = "Prenom"
    ws.Cells(3, 4).Value = "DDN": ws.Cells(3, 5).Value = "Sexe": ws.Cells(3, 6).Value = "Probleme"
    ws.Range("A3:F3").Font.Bold = True
    ws.Columns(4).NumberFormat = "@"
    r = 4
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        Dim pb As String, sx As String
        pb = ""
        sx = UCase$(Trim$(p("Sexe")))
        valeurs(sx) = valeurs(sx) + 1
        If Not modTexte.DateFrValide(Trim$(p("DDN"))) Then pb = "date de naissance": nDdn = nDdn + 1
        If Len(sx) = 0 Or (Left$(sx, 1) <> "M" And Left$(sx, 1) <> "F" And Left$(sx, 1) <> "H") Then
            pb = pb & IIf(Len(pb) > 0, " + ", "") & "sexe": nSexe = nSexe + 1
        End If
        If Len(pb) > 0 Then
            ws.Cells(r, 1).Value = p("ID"): ws.Cells(r, 2).Value = p("Nom"): ws.Cells(r, 3).Value = p("Prenom")
            ws.Cells(r, 4).Value = p("DDN"): ws.Cells(r, 5).Value = p("Sexe"): ws.Cells(r, 6).Value = pb
            r = r + 1
        End If
    Next p
    ws.Cells(2, 1).Value = (r - 4) & " fiche(s) a completer : " & nDdn & " sans date de naissance valide, " & nSexe & " sans sexe (M/F)."
    r = r + 1
    ws.Cells(r, 1).Value = "Valeurs rencontrees dans la colonne Sexe :": r = r + 1
    For Each k In valeurs.Keys
        ws.Cells(r, 1).Value = IIf(Len(k) = 0, "(vide)", k): ws.Cells(r, 2).Value = valeurs(k): r = r + 1
    Next k
    ws.Columns("A:F").AutoFit
    ws.Activate
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Sub UI_OuvrirJournal()
    modJournal.OuvrirJournal
End Sub

Public Sub UI_CalageCerfa()
    modCerfaPrint.CalageCerfa
End Sub

Public Sub UI_GrilleCerfa()
    modCerfaPrint.CalageCerfaGrille
End Sub

' Selection d'un patient (version Excel)
Public Function ChoisirPatientX() As Object
    Dim f As ufListe
    Set f = New ufListe
    f.Configurer "Choisir un patient (tapez pour filtrer)", _
                 modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS"), _
                 Array("Nom", "Prenom", "DDN", "Ville", "ID"), "120 pt;90 pt;70 pt;80 pt;40 pt"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirPatientX = f.Resultat
    Unload f
End Function
