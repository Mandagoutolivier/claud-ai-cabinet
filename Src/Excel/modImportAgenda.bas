Attribute VB_Name = "modImportAgenda"
Option Explicit
' =====================================================================
' modImportAgenda - Reprise d'un agenda papier transcrit dans un classeur
' de correction (feuille RDV : Date, Heure, Nom, Prenom, NomCorrige,
' PrenomCorrige, PatientID, Motif, DureeMin, Notes, Action).
' Regles :
'  - seules les lignes dont Action = "Importer" sont prises ;
'  - le patient est cherche par PatientID, sinon par NOM + Prenom corriges
'    (a defaut NOM + Prenom lus), insensible aux accents et a la casse ;
'  - patient introuvable : la ligne est ignoree et signalee, SAUF si
'    Action = "Importer+Creer" (fiche provisoire nom/prenom creee) ;
'  - l'agenda papier fait foi : le chevauchement n'est pas bloquant ;
'  - une ligne deja importee (colonne Resultat commencant par "OK") est
'    ignoree : le classeur peut etre rejoue sans doublon.
' Le resultat est ecrit dans la colonne Resultat du classeur.
' =====================================================================

Public Sub UI_ImporterAgenda()
    On Error GoTo Erreur
    Dim chemin As Variant, wb As Workbook, ws As Worksheet, r As Long, derniere As Long
    Dim col As Object, nImport As Long, nIgnore As Long, nCree As Long, res As String
    chemin = Application.GetOpenFilename("Classeur Excel (*.xlsx),*.xlsx", , "Agenda transcrit a importer")
    If chemin = False Then Exit Sub
    If MsgBox("Importer les rendez-vous du classeur" & vbCrLf & chemin & vbCrLf & vbCrLf & _
              "dans l'agenda du cabinet ? Seules les lignes 'Importer' sont prises ; " & _
              "les lignes deja importees (Resultat = OK) sont ignorees.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Cabinet - import d'agenda") <> vbYes Then Exit Sub

    Set wb = Workbooks.Open(CStr(chemin))
    Set ws = wb.Worksheets("RDV")
    Set col = Colonnes(ws)
    derniere = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    Application.ScreenUpdating = False
    For r = 2 To derniere
        res = ImporterLigne(ws, r, col, nCree)
        If Len(res) > 0 Then
            ws.Cells(r, col("Resultat")).Value = res
            If Left$(res, 2) = "OK" Then nImport = nImport + 1 Else nIgnore = nIgnore + 1
        End If
    Next r
    Application.ScreenUpdating = True
    wb.Save
    On Error Resume Next
    modAgendaVue.InvaliderPatients
    On Error GoTo Erreur
    MsgBox nImport & " rendez-vous importe(s), " & nCree & " fiche(s) provisoire(s) creee(s), " & _
           nIgnore & " ligne(s) non importee(s) (voir la colonne Resultat du classeur).", vbInformation, "Cabinet"
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Import interrompu : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Annule les RDV crees par un import precedent (colonne Resultat = "OK R...")
' et vide la colonne Resultat : le classeur peut alors etre corrige (duree,
' noms) et reimporte sans doublon. Les fiches provisoires creees sont
' conservees (elles seront retrouvees par nom au reimport).
Public Sub UI_AnnulerImportAgenda()
    On Error GoTo Erreur
    Dim chemin As Variant, wb As Workbook, ws As Worksheet, r As Long, derniere As Long
    Dim col As Object, res As String, id As String, n As Long, nKo As Long, annee As Long
    chemin = Application.GetOpenFilename("Classeur Excel (*.xlsx),*.xlsx", , "Classeur d'agenda deja importe")
    If chemin = False Then Exit Sub
    If MsgBox("Annuler dans l'agenda du cabinet TOUS les rendez-vous marques OK dans" & vbCrLf & chemin & vbCrLf & vbCrLf & _
              "puis vider la colonne Resultat pour permettre un nouvel import ?", _
              vbYesNo + vbExclamation + vbDefaultButton2, "Cabinet - annulation d'un import") <> vbYes Then Exit Sub
    Set wb = Workbooks.Open(CStr(chemin))
    Set ws = wb.Worksheets("RDV")
    Set col = Colonnes(ws)
    derniere = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    Application.ScreenUpdating = False
    For r = 2 To derniere
        res = Val_(ws, r, col, "Resultat")
        If Left$(res, 3) = "OK " Then
            id = Mid$(res, 4)
            If InStr(id, " ") > 0 Then id = Left$(id, InStr(id, " ") - 1)
            annee = Val(Right$(DateTexte(ws.Cells(r, col("Date")).Value), 4))
            If AnnulerRdvImporte(id, annee) Then
                n = n + 1
                ws.Cells(r, col("Resultat")).Value = ""
            Else
                nKo = nKo + 1
                ws.Cells(r, col("Resultat")).Value = "ANNULATION IMPOSSIBLE : " & res
            End If
        ElseIf Len(res) > 0 Then
            ws.Cells(r, col("Resultat")).Value = ""
        End If
    Next r
    Application.ScreenUpdating = True
    wb.Save
    On Error Resume Next
    modAgendaVue.InvaliderPatients
    On Error GoTo Erreur
    MsgBox n & " rendez-vous annule(s)" & IIf(nKo > 0, ", " & nKo & " non retrouve(s)", "") & "." & vbCrLf & _
           "Corrigez le classeur puis relancez 'Importer un agenda (xlsx)'.", vbInformation, "Cabinet"
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Annulation interrompue : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Function AnnulerRdvImporte(ByVal rdvID As String, ByVal annee As Long) As Boolean
    On Error GoTo Erreur
    Dim rdv As Object
    Set rdv = modAgenda.RdvParID(rdvID, annee)
    If rdv Is Nothing Then Exit Function
    If rdv("Statut") = "Annule" Then AnnulerRdvImporte = True: Exit Function
    modAgenda.MarquerStatut rdvID, "Annule", annee
    AnnulerRdvImporte = True
    Exit Function
Erreur:
End Function

Private Function Colonnes(ByVal ws As Worksheet) As Object
    Dim d As Object, c As Long, nom As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For c = 1 To ws.UsedRange.Columns.Count + 2
        nom = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(nom) > 0 Then d(nom) = c
    Next c
    For Each nom In Array("Date", "Heure", "Nom", "Prenom", "Action")
        If Not d.Exists(nom) Then Err.Raise vbObjectError + 640, "modImportAgenda", "Colonne absente de la feuille RDV : " & nom
    Next nom
    If Not d.Exists("Resultat") Then
        c = ws.UsedRange.Columns.Count + 1
        ws.Cells(1, c).Value = "Resultat"
        d("Resultat") = c
    End If
    Set Colonnes = d
End Function

Private Function Val_(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Object, ByVal nom As String) As String
    If col.Exists(nom) Then Val_ = Trim$(CStr(ws.Cells(r, col(nom)).Value))
End Function

Private Function ImporterLigne(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Object, ByRef nCree As Long) As String
    On Error GoTo Erreur
    Dim action As String, dateRdv As String, heure As String, nom As String, prenom As String
    Dim pat As Object, duree As String, motif As String, notes As String, id As String, dejaFait As String
    action = LCase$(Val_(ws, r, col, "Action"))
    If Left$(action, 8) <> "importer" Then Exit Function
    dejaFait = Val_(ws, r, col, "Resultat")
    If Left$(dejaFait, 2) = "OK" Then Exit Function
    dateRdv = DateTexte(ws.Cells(r, col("Date")).Value)
    heure = HeureTexte(ws.Cells(r, col("Heure")).Value)
    nom = Val_(ws, r, col, "NomCorrige"): If Len(nom) = 0 Then nom = Val_(ws, r, col, "Nom")
    prenom = Val_(ws, r, col, "PrenomCorrige"): If Len(prenom) = 0 Then prenom = Val_(ws, r, col, "Prenom")
    If Not modTexte.DateFrValide(dateRdv) Then ImporterLigne = "REFUS : date invalide (" & dateRdv & ")": Exit Function
    If Not modTexte.HeureValide(heure) Then ImporterLigne = "REFUS : heure invalide (" & heure & ")": Exit Function
    If Len(nom) = 0 Then ImporterLigne = "REFUS : nom vide": Exit Function

    Set pat = TrouverPatient(Val_(ws, r, col, "PatientID"), nom, prenom)
    If pat Is Nothing Then
        If InStr(action, "creer") > 0 Then
            Set pat = modAgenda.CreerPatientProvisoire(nom, prenom, "")
            nCree = nCree + 1
        Else
            ImporterLigne = "PATIENT INTROUVABLE : " & nom & " " & prenom & " (corrigez, ou Action = Importer+Creer)"
            Exit Function
        End If
    End If
    duree = Val_(ws, r, col, "DureeMin")
    If Val(duree) <= 0 Then duree = CStr(modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15))
    motif = Val_(ws, r, col, "Motif")
    If Len(motif) = 0 Then motif = modConfig.Config("AGENDA", "ActeDefaut", "APC")
    notes = Val_(ws, r, col, "Notes")
    If Len(notes) > 0 Then notes = "Agenda papier : " & notes Else notes = "Agenda papier"
    id = modAgenda.AjouterRdv(pat("ID"), dateRdv, heure, duree, motif, notes, True)
    ImporterLigne = "OK " & id & " (" & pat("Nom") & " " & pat("Prenom") & ")"
    Exit Function
Erreur:
    ImporterLigne = "ERREUR : " & Err.Description
End Function

Private Function TrouverPatient(ByVal patientID As String, ByVal nom As String, ByVal prenom As String) As Object
    Dim p As Object, cleNom As String, clePrenom As String, candidat As Object, n As Long
    cleNom = modTexte.Plier(nom): clePrenom = modTexte.Plier(prenom)
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If Len(patientID) > 0 Then
            If p("ID") = patientID Then Set TrouverPatient = p: Exit Function
        ElseIf modTexte.Plier(p("Nom")) = cleNom Then
            If Len(clePrenom) = 0 Or modTexte.Plier(p("Prenom")) = clePrenom Then
                Set candidat = p: n = n + 1
            ElseIf Left$(modTexte.Plier(p("Prenom")), Len(clePrenom)) = clePrenom And Len(clePrenom) >= 2 Then
                Set candidat = p: n = n + 1
            End If
        End If
    Next p
    If n = 1 Then Set TrouverPatient = candidat      ' un seul candidat : c'est lui ; plusieurs = ambigu
End Function

Private Function DateTexte(ByVal v As Variant) As String
    If VarType(v) = vbDate Then DateTexte = Format$(v, "dd/mm/yyyy") Else DateTexte = Trim$(CStr(v))
End Function

Private Function HeureTexte(ByVal v As Variant) As String
    If VarType(v) = vbDate Or (IsNumeric(v) And Val(v) < 1) Then
        HeureTexte = Format$(CDate(v), "hh:nn")
    Else
        HeureTexte = Replace(Trim$(CStr(v)), "h", ":")
    End If
End Function
