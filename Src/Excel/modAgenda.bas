Attribute VB_Name = "modAgenda"
Option Explicit
' =====================================================================
' modAgenda - Rendez-vous (poste secretaire). Fichier Agenda_AAAA.xlsx,
' feuille RDV. Statuts : Prevu / Arrive / Honore / Absent / Annule.
' =====================================================================

Public Function EntetesRdv() As Variant
    EntetesRdv = Array("ID", "PatientID", "Date", "Heure", "DureeMin", "TypeActe", _
                       "Statut", "HeureArrivee", "Notes", "DateCreation")
End Function

Public Sub AssurerAgendaAnnee(Optional ByVal annee As Long = 0)
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierAgenda(annee), "RDV", EntetesRdv()
End Sub

' Renvoie la description du RDV qui chevauche le creneau demande, "" si libre.
' Les RDV annules ne bloquent pas. rdvIDExclu permet de deplacer un RDV existant.
Public Function ConflitRdv(ByVal dateRdv As String, ByVal heure As String, _
                           ByVal dureeMin As String, _
                           Optional ByVal rdvIDExclu As String = "") As String
    Dim r As Object, duree As Long, dureeAutre As Long
    duree = Val(dureeMin)
    If duree <= 0 Then duree = modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15)
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    For Each r In modBaseIO.LireTableX(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV")
        If r("Date") = dateRdv And r("Statut") <> "Annule" Then
            If Len(rdvIDExclu) = 0 Or r("ID") <> rdvIDExclu Then
                dureeAutre = Val(r("DureeMin"))
                If dureeAutre <= 0 Then dureeAutre = modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15)
                If modTexte.IntervallesSeChevauchent(heure, duree, r("Heure"), dureeAutre) Then
                    ConflitRdv = r("Heure") & " (" & dureeAutre & " min, " & _
                                 r("TypeActe") & ", " & r("Statut") & ")"
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' forcer:=True enregistre malgre un chevauchement (double creneau assume).
Public Function AjouterRdv(ByVal patientID As String, ByVal dateRdv As String, _
                           ByVal heure As String, ByVal dureeMin As String, _
                           ByVal typeActe As String, ByVal notes As String, _
                           Optional ByVal forcer As Boolean = False) As String
    Dim d As Object, conflit As String
    If Len(Trim$(patientID)) = 0 Then
        Err.Raise vbObjectError + 620, "modAgenda", "Aucun patient : le rendez-vous n'a pas ete enregistre."
    End If
    If Not modTexte.DateFrValide(dateRdv) Then
        Err.Raise vbObjectError + 621, "modAgenda", _
            "Date de rendez-vous invalide : " & dateRdv & " (attendu jj/mm/aaaa, date existante)."
    End If
    If Not modTexte.HeureValide(heure) Then
        Err.Raise vbObjectError + 622, "modAgenda", _
            "Heure de rendez-vous invalide : " & heure & " (attendu hh:mm entre 00:00 et 23:59)."
    End If
    If Val(dureeMin) <= 0 Then
        Err.Raise vbObjectError + 623, "modAgenda", "Duree de rendez-vous invalide : " & dureeMin & " minutes."
    End If
    If Not forcer Then
        conflit = ConflitRdv(dateRdv, heure, dureeMin)
        If Len(conflit) > 0 Then
            Err.Raise vbObjectError + 624, "modAgenda", _
                "Ce creneau chevauche un rendez-vous deja prevu le " & dateRdv & " a " & conflit & "."
        End If
    End If
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = patientID
    d("Date") = dateRdv
    d("Heure") = heure
    d("DureeMin") = dureeMin
    d("TypeActe") = typeActe
    d("Statut") = "Prevu"
    d("Notes") = notes
    d("DateCreation") = Format$(Now, "dd/mm/yyyy hh:nn")
    AjouterRdv = modBaseIO.AjouterLigne(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV", d, "R")
End Function

Public Sub MarquerStatut(ByVal rdvID As String, ByVal statut As String, _
                         Optional ByVal annee As Long = 0)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Statut") = statut
    If statut = "Arrive" Then d("HeureArrivee") = Format$(Now, "hh:nn")
    modBaseIO.ModifierLigne modConfig.FichierAgenda(annee), "RDV", "ID", rdvID, d
End Sub

' RDV d'une date (defaut aujourd'hui), tries par heure
Public Function RdvDuJour(Optional ByVal dateRdv As String = "") As Collection
    Dim tous As Collection, jour As Collection, r As Object
    Dim pats As Object, p As Object
    If Len(dateRdv) = 0 Then dateRdv = Format$(Date, "dd/mm/yyyy")
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    Set tous = modBaseIO.LireTableX(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV")
    ' index patients pour l'affichage
    Set pats = CreateObject("Scripting.Dictionary")
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        Set pats(p("ID")) = p
    Next p
    Set jour = New Collection
    For Each r In tous
        If r("Date") = dateRdv Then
            If pats.Exists(r("PatientID")) Then
                r("Nom") = pats(r("PatientID"))("Nom")
                r("Prenom") = pats(r("PatientID"))("Prenom")
            Else
                r("Nom") = "?"
                r("Prenom") = ""
            End If
            InsererParHeure jour, r
        End If
    Next r
    Set RdvDuJour = jour
End Function

Private Sub InsererParHeure(ByVal col As Collection, ByVal r As Object)
    Dim i As Long
    For i = 1 To col.Count
        If col(i)("Heure") > r("Heure") Then
            col.Add r, , i
            Exit Sub
        End If
    Next i
    col.Add r
End Sub

Private Function AnneeDeDate(ByVal dateTexte As String) As Long
    Dim parties() As String
    parties = Split(dateTexte, "/")
    If UBound(parties) = 2 Then AnneeDeDate = Val(parties(2)) Else AnneeDeDate = Year(Date)
End Function
