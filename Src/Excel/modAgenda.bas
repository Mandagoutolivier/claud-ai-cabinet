Attribute VB_Name = "modAgenda"
Option Explicit
' =====================================================================
' modAgenda - Rendez-vous (poste secretaire). Fichier Agenda_AAAA.xlsx,
' feuille RDV.
' Statuts : Prevu / Arrive / Honore / Absent / Annule / Bloque.
'  - Bloque = INDISPONIBILITE (conges, formation, reunion...) : une ligne
'    sans patient (PatientID vide, TypeActe = INDISPO, motif dans Notes)
'    qui occupe le creneau comme un rendez-vous.
'  - ConsultationID : pose quand la seance est enregistree au journal
'    (lien rendez-vous -> consultation, audit 05/09/2026).
' =====================================================================

Public Const STATUT_BLOQUE As String = "Bloque"
Public Const TYPE_INDISPO As String = "INDISPO"

Public Function EntetesRdv() As Variant
    EntetesRdv = Array("ID", "PatientID", "Date", "Heure", "DureeMin", "TypeActe", _
                       "Statut", "HeureArrivee", "Notes", "DateCreation", "ConsultationID")
End Function

Public Sub AssurerAgendaAnnee(Optional ByVal annee As Long = 0)
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierAgenda(annee), "RDV", EntetesRdv()
    ' agendas crees par une version anterieure : completer les colonnes
    modBaseIO.AssurerColonnes modConfig.FichierAgenda(annee), "RDV", EntetesRdv()
End Sub

Public Function EstBlocage(ByVal r As Object) As Boolean
    EstBlocage = (r("Statut") = STATUT_BLOQUE Or r("TypeActe") = TYPE_INDISPO)
End Function

' Renvoie la description du RDV (ou de l'indisponibilite) qui chevauche le
' creneau demande, "" si libre. Les RDV annules ne bloquent pas.
' rdvIDExclu permet de deplacer un RDV existant sans qu'il se gene lui-meme.
Public Function ConflitRdv(ByVal dateRdv As String, ByVal heure As String, _
                           ByVal dureeMin As String, _
                           Optional ByVal rdvIDExclu As String = "") As String
    Dim r As Object, duree As Long, dureeAutre As Long, debut As Long, debutAutre As Long
    duree = Val(dureeMin)
    If duree <= 0 Then duree = modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15)
    ' IntervallesSeChevauchent travaille en MINUTES depuis minuit, pas en hh:mm
    debut = modTexte.MinutesDepuisMinuit(heure)
    If debut < 0 Then Exit Function
    AssurerAgendaAnnee AnneeDeDate(dateRdv)
    For Each r In modBaseIO.LireTableX(modConfig.FichierAgenda(AnneeDeDate(dateRdv)), "RDV")
        If r("Date") = dateRdv And r("Statut") <> "Annule" Then
            If Len(rdvIDExclu) = 0 Or r("ID") <> rdvIDExclu Then
                dureeAutre = Val(r("DureeMin"))
                If dureeAutre <= 0 Then dureeAutre = modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15)
                debutAutre = modTexte.MinutesDepuisMinuit(r("Heure"))
                If debutAutre >= 0 And modTexte.IntervallesSeChevauchent(debut, duree, debutAutre, dureeAutre) Then
                    If EstBlocage(r) Then
                        ConflitRdv = r("Heure") & " (INDISPONIBLE " & dureeAutre & " min : " & r("Notes") & ")"
                    Else
                        ConflitRdv = r("Heure") & " (" & dureeAutre & " min, " & _
                                     r("TypeActe") & ", " & r("Statut") & ")"
                    End If
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' Validations communes a l'ajout et a la modification
Private Sub ValiderCreneau(ByVal dateRdv As String, ByVal heure As String, ByVal dureeMin As String)
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
End Sub

' forcer:=True enregistre malgre un chevauchement (double creneau assume).
Public Function AjouterRdv(ByVal patientID As String, ByVal dateRdv As String, _
                           ByVal heure As String, ByVal dureeMin As String, _
                           ByVal typeActe As String, ByVal notes As String, _
                           Optional ByVal forcer As Boolean = False) As String
    Dim d As Object, conflit As String
    If Len(Trim$(patientID)) = 0 Then
        Err.Raise vbObjectError + 620, "modAgenda", "Aucun patient : le rendez-vous n'a pas ete enregistre."
    End If
    ValiderCreneau dateRdv, heure, dureeMin
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

' Deplacement / modification d'un RDV existant. Le RDV garde son ID et son
' statut si l'annee ne change pas ; sinon (changement d'annee = autre
' classeur) l'ancien est annule et un nouveau est cree : l'ID renvoye est
' celui du RDV en vigueur.
Public Function ModifierRdv(ByVal rdvID As String, ByVal anneeOrigine As Long, _
                            ByVal dateRdv As String, ByVal heure As String, _
                            ByVal dureeMin As String, ByVal typeActe As String, _
                            ByVal notes As String, Optional ByVal forcer As Boolean = False) As String
    Dim d As Object, conflit As String, ancien As Object
    ValiderCreneau dateRdv, heure, dureeMin
    Set ancien = RdvParID(rdvID, anneeOrigine)
    If ancien Is Nothing Then
        Err.Raise vbObjectError + 625, "modAgenda", "Rendez-vous introuvable : " & rdvID
    End If
    If Not forcer Then
        conflit = ConflitRdv(dateRdv, heure, dureeMin, rdvID)
        If Len(conflit) > 0 Then
            Err.Raise vbObjectError + 624, "modAgenda", _
                "Ce creneau chevauche un rendez-vous deja prevu le " & dateRdv & " a " & conflit & "."
        End If
    End If
    If AnneeDeDate(dateRdv) = anneeOrigine Then
        Set d = CreateObject("Scripting.Dictionary")
        d("Date") = dateRdv
        d("Heure") = heure
        d("DureeMin") = dureeMin
        d("TypeActe") = typeActe
        d("Notes") = notes
        modBaseIO.ModifierLigne modConfig.FichierAgenda(anneeOrigine), "RDV", "ID", rdvID, d
        ModifierRdv = rdvID
    Else
        ModifierRdv = AjouterRdv(ancien("PatientID"), dateRdv, heure, dureeMin, typeActe, notes, True)
        MarquerStatut rdvID, "Annule", anneeOrigine
    End If
End Function

' Indisponibilite sur une periode : chaque jour de dateDebut a dateFin,
' de heureDebut a heureFin (vides = journee entiere de l'agenda).
' Renvoie le nombre de jours bloques.
Public Function BloquerPeriode(ByVal dateDebut As String, ByVal dateFin As String, _
                               ByVal heureDebut As String, ByVal heureFin As String, _
                               ByVal motif As String) As Long
    Dim d1 As Date, d2 As Date, j As Date, d As Object, duree As Long, n As Long
    If Not modTexte.DateFrValide(dateDebut) Then Err.Raise vbObjectError + 621, "modAgenda", "Date de debut invalide : " & dateDebut
    If Len(Trim$(dateFin)) = 0 Then dateFin = dateDebut
    If Not modTexte.DateFrValide(dateFin) Then Err.Raise vbObjectError + 621, "modAgenda", "Date de fin invalide : " & dateFin
    If Len(Trim$(heureDebut)) = 0 Then heureDebut = modConfig.Config("AGENDA", "HeureDebut", "08:00")
    If Len(Trim$(heureFin)) = 0 Then heureFin = modConfig.Config("AGENDA", "HeureFin", "19:00")
    If Not modTexte.HeureValide(heureDebut) Or Not modTexte.HeureValide(heureFin) Then
        Err.Raise vbObjectError + 622, "modAgenda", "Heures invalides : " & heureDebut & " - " & heureFin
    End If
    duree = modTexte.MinutesDepuisMinuit(heureFin) - modTexte.MinutesDepuisMinuit(heureDebut)
    If duree <= 0 Then Err.Raise vbObjectError + 623, "modAgenda", "L'heure de fin doit etre apres l'heure de debut."
    d1 = modTexte.DateFr(dateDebut): d2 = modTexte.DateFr(dateFin)
    If d2 < d1 Then Err.Raise vbObjectError + 621, "modAgenda", "La date de fin est avant la date de debut."
    If d2 - d1 > 366 Then Err.Raise vbObjectError + 621, "modAgenda", "Periode trop longue (plus d'un an)."
    If Len(Trim$(motif)) = 0 Then motif = "Indisponible"
    For j = d1 To d2
        AssurerAgendaAnnee Year(j)
        Set d = CreateObject("Scripting.Dictionary")
        d("PatientID") = ""
        d("Date") = Format$(j, "dd/mm/yyyy")
        d("Heure") = heureDebut
        d("DureeMin") = CStr(duree)
        d("TypeActe") = TYPE_INDISPO
        d("Statut") = STATUT_BLOQUE
        d("Notes") = Trim$(motif)
        d("DateCreation") = Format$(Now, "dd/mm/yyyy hh:nn")
        modBaseIO.AjouterLigne modConfig.FichierAgenda(Year(j)), "RDV", d, "B"
        n = n + 1
    Next j
    BloquerPeriode = n
End Function

Public Sub MarquerStatut(ByVal rdvID As String, ByVal statut As String, _
                         Optional ByVal annee As Long = 0)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("Statut") = statut
    If statut = "Arrive" Then d("HeureArrivee") = Format$(Now, "hh:nn")
    modBaseIO.ModifierLigne modConfig.FichierAgenda(annee), "RDV", "ID", rdvID, d
End Sub

Public Function RdvParID(ByVal rdvID As String, Optional ByVal annee As Long = 0) As Object
    Dim r As Object
    AssurerAgendaAnnee annee
    For Each r In modBaseIO.LireTableX(modConfig.FichierAgenda(annee), "RDV")
        If r("ID") = rdvID Then Set RdvParID = r: Exit Function
    Next r
End Function

' --- lien rendez-vous <-> consultation ---------------------------------
' RDV d'un patient un jour donne (hors annules et indisponibilites), "" si
' aucun. Sert a rattacher la seance enregistree au journal.
Public Function RdvPourSeance(ByVal patientID As String, ByVal dateActe As String) As String
    Dim r As Object
    If Len(patientID) = 0 Or Not modTexte.DateFrValide(dateActe) Then Exit Function
    On Error Resume Next
    AssurerAgendaAnnee AnneeDeDate(dateActe)
    For Each r In modBaseIO.LireTableX(modConfig.FichierAgenda(AnneeDeDate(dateActe)), "RDV")
        If r("PatientID") = patientID And r("Date") = dateActe And r("Statut") <> "Annule" _
           And Not EstBlocage(r) Then
            RdvPourSeance = r("ID")
            If r("Statut") <> "Honore" Then Exit Function   ' le premier non encore honore
        End If
    Next r
End Function

' Pose ConsultationID sur le RDV et le passe a Honore (le patient a ete vu).
Public Sub LierConsultation(ByVal rdvID As String, ByVal consultationID As String, ByVal annee As Long)
    Dim d As Object
    If Len(rdvID) = 0 Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary")
    d("ConsultationID") = consultationID
    d("Statut") = "Honore"
    modBaseIO.ModifierLigne modConfig.FichierAgenda(annee), "RDV", "ID", rdvID, d
End Sub

' RDV d'une date (defaut aujourd'hui), tries par heure, sans les
' indisponibilites (liste d'arrivee des patients)
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
        If r("Date") = dateRdv And Not EstBlocage(r) Then
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
