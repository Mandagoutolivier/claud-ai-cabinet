Attribute VB_Name = "modAgendaVue"
Option Explicit
' =====================================================================
' modAgendaVue - Agenda (feuille "Agenda"), deux vues :
'  - SEMAINE : une colonne par jour, une ligne par creneau (config [AGENDA]).
'    Double-clic : creneau libre -> prise de RDV ; creneau occupe -> actions
'    (arrive, absent, deplacer/modifier, annuler, fiche) ; indisponibilite
'    -> lever le blocage.
'  - MOIS : un calendrier, une case par jour avec le nombre de RDV et les
'    indisponibilites. Double-clic sur un jour -> vue semaine de ce jour.
' Navigation : semaine, mois, aujourd'hui, aller a une date.
' Les donnees restent dans Agenda_AAAA.xlsx (modAgenda) ; la grille est
' un simple rendu, rafraichi apres chaque action et, en arriere-plan, quand
' le fichier a change (autre poste) : [AGENDA] RafraichirSecondes.
' =====================================================================

Private mLundi As Date                 ' lundi de la semaine affichee
Private mMois As Date                  ' 1er du mois affiche (vue mois)
Private mVueMois As Boolean
Private mPatientPreselect As Object    ' patient pour prise de RDV directe (dict) ou Nothing
Private mCarte As Object               ' "ligne|col" -> Collection de RDV (dict) du creneau
Private mCarteJours As Object          ' vue mois : "ligne|col" -> date (Double)
Private mPatients As Object            ' ID -> dict patient (cache du rendu)
Private mProchainRafraichissement As Date
Private mRafraichissementActif As Boolean
Private mEmpreinte As String           ' dates de modification des fichiers agenda affiches

Private Const LIG_TITRE As Long = 2
Private Const LIG_ENTETE As Long = 3
Private Const LIG_DEBUT As Long = 4
Private Const COL_HEURE As Long = 1
Private Const COL_PREMIER_JOUR As Long = 2

' --- parametres -------------------------------------------------------
Private Function NbJours() As Long
    NbJours = CLng(modConfig.ConfigNum("AGENDA", "Jours", 5))
    If NbJours < 1 Then NbJours = 1
    If NbJours > 7 Then NbJours = 7
End Function

Private Function PasMinutes() As Long
    PasMinutes = CLng(modConfig.ConfigNum("AGENDA", "PasMinutes", 15))
    If PasMinutes < 5 Then PasMinutes = 5
End Function

Private Function HeureDeLigne(ByVal ligne As Long) As String
    Dim minutes As Long
    minutes = MinutesDebut() + (ligne - LIG_DEBUT) * PasMinutes()
    If ligne < LIG_DEBUT Or minutes >= MinutesFin() Then HeureDeLigne = "" Else HeureDeLigne = TexteHeure(minutes)
End Function

Private Function MinutesDebut() As Long
    MinutesDebut = MinutesDe(modConfig.Config("AGENDA", "HeureDebut", "08:00"))
End Function

Private Function MinutesFin() As Long
    MinutesFin = MinutesDe(modConfig.Config("AGENDA", "HeureFin", "19:00"))
End Function

Private Function MinutesDe(ByVal hhmm As String) As Long
    Dim p() As String
    p = Split(Replace(Trim$(hhmm), "h", ":"), ":")
    If UBound(p) >= 1 Then MinutesDe = Val(p(0)) * 60 + Val(p(1)) Else MinutesDe = Val(p(0)) * 60
End Function

Private Function TexteHeure(ByVal minutes As Long) As String
    TexteHeure = Format$(minutes \ 60, "00") & ":" & Format$(minutes Mod 60, "00")
End Function

Private Function LundiDe(ByVal d As Date) As Date
    LundiDe = d - Weekday(d, vbMonday) + 1
End Function

Private Function DateDe(ByVal texte As String) As Date
    Dim p() As String
    p = Split(Trim$(texte), "/")
    If UBound(p) = 2 Then DateDe = DateSerial(Val(p(2)), Val(p(1)), Val(p(0))) Else DateDe = 0
End Function

' --- navigation -------------------------------------------------------
Public Sub AfficherAgenda(Optional ByVal dateRef As Date = 0, Optional ByVal pat As Object = Nothing)
    If dateRef = 0 Then dateRef = Date
    Set mPatientPreselect = pat
    mVueMois = False
    mLundi = LundiDe(dateRef)
    mMois = DateSerial(Year(dateRef), Month(dateRef), 1)
    Rendre
    DemarrerRafraichissement
End Sub

Public Sub AgendaSemainePrecedente()
    Initialiser
    If mVueMois Then AgendaMoisPrecedent: Exit Sub
    mLundi = mLundi - 7
    Rendre
End Sub

Public Sub AgendaSemaineSuivante()
    Initialiser
    If mVueMois Then AgendaMoisSuivant: Exit Sub
    mLundi = mLundi + 7
    Rendre
End Sub

' Mois precedent / suivant : en vue semaine, on garde le meme jour de la
' semaine un mois plus tot/tard (lundi de la semaine correspondante).
Public Sub AgendaMoisPrecedent()
    Initialiser
    If mVueMois Then
        mMois = DateAdd("m", -1, mMois)
    Else
        mLundi = LundiDe(DateAdd("m", -1, mLundi))
    End If
    Rendre
End Sub

Public Sub AgendaMoisSuivant()
    Initialiser
    If mVueMois Then
        mMois = DateAdd("m", 1, mMois)
    Else
        mLundi = LundiDe(DateAdd("m", 1, mLundi))
    End If
    Rendre
End Sub

Public Sub AgendaAujourdhui()
    mLundi = LundiDe(Date)
    mMois = DateSerial(Year(Date), Month(Date), 1)
    Rendre
End Sub

' Bascule semaine <-> mois en restant sur la meme periode
Public Sub AgendaVueMois()
    Initialiser
    If mVueMois Then
        mVueMois = False
        If mLundi < mMois Or mLundi > DateAdd("m", 1, mMois) Then mLundi = LundiDe(mMois)
    Else
        mVueMois = True
        mMois = DateSerial(Year(mLundi + 3), Month(mLundi + 3), 1)   ' mois du milieu de semaine
    End If
    Rendre
End Sub

' Aller directement a une date (jj/mm/aaaa, ou +N / -N jours)
Public Sub AgendaAllerA()
    Dim s As String, d As Date
    Initialiser
    s = Trim$(InputBox("Date a afficher (jj/mm/aaaa)," & vbCrLf & _
                       "ou +N / -N pour N jours a partir d'aujourd'hui :", _
                       "Aller a une date", Format$(Date, "dd/mm/yyyy")))
    If Len(s) = 0 Then Exit Sub
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then
        d = Date + Val(s)
    ElseIf modTexte.DateFrValide(s) Then
        d = modTexte.DateFr(s)
    Else
        MsgBox "Date invalide : " & s, vbExclamation, "Cabinet"
        Exit Sub
    End If
    mLundi = LundiDe(d)
    mMois = DateSerial(Year(d), Month(d), 1)
    Rendre
End Sub

' Indisponibilite (conges, formation, reunion...) sur une periode
Public Sub AgendaBloquer()
    On Error GoTo Erreur
    Dim d1 As String, d2 As String, h1 As String, h2 As String, motif As String, n As Long
    Initialiser
    d1 = Trim$(InputBox("Premier jour indisponible (jj/mm/aaaa) :", "Bloquer des creneaux", _
                        Format$(IIf(mVueMois, mMois, mLundi), "dd/mm/yyyy")))
    If Len(d1) = 0 Then Exit Sub
    d2 = Trim$(InputBox("Dernier jour indisponible (jj/mm/aaaa) - vide = le meme jour :", "Bloquer des creneaux", d1))
    h1 = Trim$(InputBox("Heure de debut (hh:mm) - vide = journee entiere :", "Bloquer des creneaux", ""))
    If Len(h1) > 0 Then
        h2 = Trim$(InputBox("Heure de fin (hh:mm) :", "Bloquer des creneaux", modConfig.Config("AGENDA", "HeureFin", "19:00")))
        If Len(h2) = 0 Then Exit Sub
    End If
    motif = Trim$(InputBox("Motif (affiche dans l'agenda) :", "Bloquer des creneaux", "Indisponible"))
    n = modAgenda.BloquerPeriode(d1, d2, h1, h2, motif)
    MsgBox n & " jour(s) bloque(s) : " & motif, vbInformation, "Cabinet"
    Rendre
    Exit Sub
Erreur:
    MsgBox "Blocage impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Public Sub AgendaRetourAccueil()
    Set mPatientPreselect = Nothing
    ArreterRafraichissement
    ThisWorkbook.Worksheets("Accueil").Activate
End Sub

Private Sub Initialiser()
    If mLundi = 0 Then mLundi = LundiDe(Date)
    If mMois = 0 Then mMois = DateSerial(Year(Date), Month(Date), 1)
End Sub

Private Function Feuille() As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets("Agenda")
    If Feuille Is Nothing Then
        Set Feuille = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Feuille.Name = "Agenda"
    End If
End Function

' --- rafraichissement en arriere-plan (autre poste) ---------------------
' Ne redessine que si un fichier agenda affiche a change, et sans voler la
' selection (silencieux).
Public Sub DemarrerRafraichissement()
    ArreterRafraichissement
    Dim secondes As Long
    secondes = CLng(modConfig.ConfigNum("AGENDA", "RafraichirSecondes", 60))
    If secondes <= 0 Then Exit Sub
    If secondes < 15 Then secondes = 15
    mProchainRafraichissement = Now + TimeSerial(0, 0, secondes)
    mRafraichissementActif = True
    Application.OnTime mProchainRafraichissement, NomMacroRafraichissement()
End Sub

Public Sub ArreterRafraichissement()
    On Error Resume Next
    If mRafraichissementActif Then
        Application.OnTime mProchainRafraichissement, NomMacroRafraichissement(), , False
    End If
    mRafraichissementActif = False
End Sub

Private Function NomMacroRafraichissement() As String
    NomMacroRafraichissement = "'" & ThisWorkbook.Name & "'!modAgendaVue.RafraichirSiModifie"
End Function

Public Sub RafraichirSiModifie()
    On Error Resume Next
    mRafraichissementActif = False
    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> "Agenda" Then Exit Sub        ' on a quitte l'agenda : on s'arrete
    If Empreinte() <> mEmpreinte Then Rendre True
    DemarrerRafraichissement
End Sub

' Dates de modification des classeurs agenda concernes par l'affichage
Private Function Empreinte() As String
    Dim a As Variant, f As String
    On Error Resume Next
    For Each a In AnneesAffichees()
        f = modConfig.FichierAgenda(CLng(a))
        If Len(Dir$(f)) > 0 Then Empreinte = Empreinte & a & "=" & FileDateTime(f) & ";"
    Next a
End Function

Private Function AnneesAffichees() As Variant
    Dim d1 As Date, d2 As Date
    If mVueMois Then
        d1 = mMois: d2 = DateAdd("m", 1, mMois) - 1
    Else
        d1 = mLundi: d2 = mLundi + NbJours() - 1
    End If
    If Year(d1) = Year(d2) Then AnneesAffichees = Array(Year(d1)) Else AnneesAffichees = Array(Year(d1), Year(d2))
End Function

' --- rendu --------------------------------------------------------------
Public Sub Rendre(Optional ByVal silencieux As Boolean = False)
    On Error GoTo Erreur
    Initialiser
    Application.ScreenUpdating = False
    ChargerPatients
    If mVueMois Then RendreMois silencieux Else RendreSemaine silencieux
    mEmpreinte = Empreinte()
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    If Not silencieux Then MsgBox "Affichage de l'agenda impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Private Sub Nettoyer(ByVal ws As Worksheet)
    ws.Range(ws.Cells(LIG_TITRE, 1), ws.Cells(ws.Rows.Count, COL_PREMIER_JOUR + 7)).Clear
    ws.Range(ws.Cells(LIG_TITRE, 1), ws.Cells(ws.Rows.Count, COL_PREMIER_JOUR + 7)).RowHeight = 15
End Sub

Private Sub RendreSemaine(ByVal silencieux As Boolean)
    Dim ws As Worksheet, j As Long, r As Long, minutes As Long, nbLignes As Long
    Dim rdv As Object, jours As Variant
    jours = Array("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")
    Set ws = Feuille()
    Nettoyer ws
    ws.Cells(LIG_TITRE, 1).Value = "Semaine du " & Format$(mLundi, "dd/mm/yyyy") & " au " & _
                                   Format$(mLundi + NbJours() - 1, "dd/mm/yyyy")
    ws.Cells(LIG_TITRE, 1).Font.Bold = True
    ws.Cells(LIG_TITRE, 1).Font.Size = 12
    If Not mPatientPreselect Is Nothing Then
        ws.Cells(LIG_TITRE, 4).Value = "RDV pour " & mPatientPreselect("Prenom") & " " & mPatientPreselect("Nom") & _
                                       " : double-cliquez un creneau libre"
        ws.Cells(LIG_TITRE, 4).Font.Color = RGB(192, 0, 0)
        ws.Cells(LIG_TITRE, 4).Font.Bold = True
    Else
        ws.Cells(LIG_TITRE, 4).Value = "Double-clic : creneau libre = prendre RDV | creneau occupe = arrive / absent / deplacer / annuler / fiche"
        ws.Cells(LIG_TITRE, 4).Font.Color = RGB(90, 90, 90)
    End If

    ws.Cells(LIG_ENTETE, COL_HEURE).Value = "Heure"
    For j = 0 To NbJours() - 1
        ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + j).Value = jours(j) & " " & Format$(mLundi + j, "dd/mm")
        If mLundi + j = Date Then ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + j).Interior.Color = RGB(255, 235, 156)
    Next j
    With ws.Range(ws.Cells(LIG_ENTETE, 1), ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + NbJours() - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns(COL_HEURE).NumberFormat = "@"
    nbLignes = 0
    For minutes = MinutesDebut() To MinutesFin() - PasMinutes() Step PasMinutes()
        r = LIG_DEBUT + nbLignes
        ws.Cells(r, COL_HEURE).Value = TexteHeure(minutes)
        If (minutes Mod 60) <> 0 Then ws.Cells(r, COL_HEURE).Font.Color = RGB(120, 120, 120)
        ws.Cells(r, COL_HEURE).HorizontalAlignment = xlCenter
        ws.Cells(r, COL_HEURE).Font.Bold = True
        With ws.Range(ws.Cells(r, COL_PREMIER_JOUR), ws.Cells(r, COL_PREMIER_JOUR + NbJours() - 1))
            .Interior.Color = RGB(235, 250, 235)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(200, 200, 200)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        nbLignes = nbLignes + 1
    Next minutes
    ws.Range(ws.Cells(LIG_DEBUT, 1), ws.Cells(LIG_DEBUT + nbLignes - 1, 1)).Borders.LineStyle = xlContinuous

    Set mCarte = CreateObject("Scripting.Dictionary")
    Set mCarteJours = Nothing
    For Each rdv In RdvPeriode(mLundi, mLundi + NbJours() - 1)
        PlacerRdv ws, rdv
    Next rdv

    ws.Columns(COL_HEURE).ColumnWidth = 7
    For j = 0 To 6
        ws.Columns(COL_PREMIER_JOUR + j).ColumnWidth = IIf(j < NbJours(), 26, 8)
    Next j
    ws.Rows(LIG_DEBUT & ":" & (LIG_DEBUT + nbLignes - 1)).RowHeight = IIf(PasMinutes() < 30, 16, 30)
    If Not silencieux Then
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.FreezePanes = False
        ws.Cells(LIG_DEBUT, COL_PREMIER_JOUR).Select
        ActiveWindow.FreezePanes = True
        ws.Cells(LIG_DEBUT, COL_PREMIER_JOUR).Select
    End If
End Sub

' Vue mois : 7 colonnes (lundi -> dimanche), une ligne par semaine.
' Chaque case : numero du jour, nombre de RDV, indisponibilites.
Private Sub RendreMois(ByVal silencieux As Boolean)
    Dim ws As Worksheet, premier As Date, dernier As Date, j As Date, jours As Variant
    Dim ligne As Long, col As Long, rdv As Object, parJour As Object, cle As String
    Dim blocs As Object, texte As String, nbSem As Long, d As Date, nomsMois As Variant
    jours = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche")
    nomsMois = Array("janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", _
                     "aout", "septembre", "octobre", "novembre", "decembre")
    Set ws = Feuille()
    Nettoyer ws
    premier = mMois
    dernier = DateAdd("m", 1, mMois) - 1
    ws.Cells(LIG_TITRE, 1).Value = UCase$(Left$(nomsMois(Month(mMois) - 1), 1)) & _
                                   Mid$(nomsMois(Month(mMois) - 1), 2) & " " & Year(mMois)
    ws.Cells(LIG_TITRE, 1).Font.Bold = True
    ws.Cells(LIG_TITRE, 1).Font.Size = 14
    ws.Cells(LIG_TITRE, 4).Value = "Double-clic sur un jour : ouvrir la semaine | Vue mois / semaine pour revenir"
    ws.Cells(LIG_TITRE, 4).Font.Color = RGB(90, 90, 90)

    For col = 0 To 6
        ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + col).Value = jours(col)
    Next col
    With ws.Range(ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR), ws.Cells(LIG_ENTETE, COL_PREMIER_JOUR + 6))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' comptage par jour
    Set parJour = CreateObject("Scripting.Dictionary")
    Set blocs = CreateObject("Scripting.Dictionary")
    For Each rdv In RdvPeriode(premier, dernier)
        If rdv("Statut") <> "Annule" Then
            cle = rdv("Date")
            If modAgenda.EstBlocage(rdv) Then
                If blocs.Exists(cle) Then
                    If InStr(blocs(cle), rdv("Notes")) = 0 Then blocs(cle) = blocs(cle) & ", " & rdv("Notes")
                Else
                    blocs(cle) = rdv("Notes")
                End If
            Else
                parJour(cle) = parJour(cle) + 1
            End If
        End If
    Next rdv

    Set mCarteJours = CreateObject("Scripting.Dictionary")
    Set mCarte = Nothing
    j = LundiDe(premier)
    ligne = LIG_DEBUT
    nbSem = 0
    Do While j <= dernier
        For col = 0 To 6
            d = j + col
            With ws.Cells(ligne, COL_PREMIER_JOUR + col)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(200, 200, 200)
                .VerticalAlignment = xlTop
                .WrapText = True
                If d < premier Or d > dernier Then
                    .Interior.Color = RGB(245, 245, 245)
                Else
                    cle = Format$(d, "dd/mm/yyyy")
                    texte = Day(d)
                    If parJour.Exists(cle) Then texte = texte & vbLf & parJour(cle) & " RDV"
                    If blocs.Exists(cle) Then texte = texte & vbLf & "INDISPO : " & blocs(cle)
                    .Value = texte
                    .Font.Size = 10
                    If col >= NbJours() Then
                        .Interior.Color = RGB(240, 240, 240)          ' hors jours ouvres
                    ElseIf blocs.Exists(cle) Then
                        .Interior.Color = RGB(217, 217, 217)
                    ElseIf parJour.Exists(cle) Then
                        .Interior.Color = RGB(255, 242, 180)
                    Else
                        .Interior.Color = RGB(235, 250, 235)
                    End If
                    If d = Date Then .Font.Bold = True: .Borders.Color = RGB(192, 0, 0): .Borders.Weight = xlMedium
                    mCarteJours(ligne & "|" & (COL_PREMIER_JOUR + col)) = CDbl(d)
                End If
            End With
        Next col
        ws.Rows(ligne).RowHeight = 48
        ligne = ligne + 1
        nbSem = nbSem + 1
        j = j + 7
    Loop
    ws.Columns(COL_HEURE).ColumnWidth = 2
    For col = 0 To 6
        ws.Columns(COL_PREMIER_JOUR + col).ColumnWidth = 22
    Next col
    If Not silencieux Then
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Cells(LIG_DEBUT, COL_PREMIER_JOUR).Select
    End If
End Sub

Private Sub ChargerPatients()
    Dim p As Object
    Set mPatients = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        Set mPatients(p("ID")) = p
    Next p
End Sub

' RDV entre deux dates (les deux annees si la periode chevauche)
Private Function RdvPeriode(ByVal d1 As Date, ByVal d2 As Date) As Collection
    Dim col As Collection, r As Object, d As Date, annees As Object, a As Variant, tous As Collection
    Set col = New Collection
    Set annees = CreateObject("Scripting.Dictionary")
    annees(Year(d1)) = 1
    annees(Year(d2)) = 1
    For Each a In annees.Keys
        modAgenda.AssurerAgendaAnnee CLng(a)
        On Error Resume Next
        Set tous = Nothing
        Set tous = modBaseIO.LireTableX(modConfig.FichierAgenda(CLng(a)), "RDV")
        On Error GoTo 0
        If Not tous Is Nothing Then
            For Each r In tous
                d = DateDe(r("Date"))
                If d >= d1 And d <= d2 Then col.Add r
            Next r
        End If
    Next a
    Set RdvPeriode = col
End Function

Private Sub PlacerRdv(ByVal ws As Worksheet, ByVal rdv As Object)
    Dim d As Date, minutes As Long, ligne As Long, col As Long, cle As String
    Dim nom As String, texte As String, nbCreneaux As Long, k As Long, statut As String
    Dim blocage As Boolean, derniereLigne As Long
    If rdv("Statut") = "Annule" Then Exit Sub
    d = DateDe(rdv("Date"))
    minutes = MinutesDe(rdv("Heure"))
    blocage = modAgenda.EstBlocage(rdv)
    ' une indisponibilite qui commence avant l'ouverture est ramenee a l'ouverture
    If blocage And minutes < MinutesDebut() Then minutes = MinutesDebut()
    If minutes < MinutesDebut() Or minutes >= MinutesFin() Then Exit Sub
    col = COL_PREMIER_JOUR + (d - mLundi)
    If col < COL_PREMIER_JOUR Or col > COL_PREMIER_JOUR + NbJours() - 1 Then Exit Sub
    ligne = LIG_DEBUT + (minutes - MinutesDebut()) \ PasMinutes()
    derniereLigne = LIG_DEBUT + (MinutesFin() - MinutesDebut()) \ PasMinutes() - 1
    cle = ligne & "|" & col
    statut = rdv("Statut")
    If blocage Then
        nom = "INDISPONIBLE"
        texte = nom & " - " & rdv("Notes")
    Else
        If mPatients.Exists(rdv("PatientID")) Then
            nom = mPatients(rdv("PatientID"))("Nom") & " " & mPatients(rdv("PatientID"))("Prenom")
        Else
            nom = rdv("PatientID")
        End If
        texte = nom
        If Len(rdv("TypeActe")) > 0 Then texte = texte & " - " & rdv("TypeActe")
        If statut = "Arrive" Then texte = texte & " (arrive " & rdv("HeureArrivee") & ")"
        If statut = "Absent" Then texte = texte & " (ABSENT)"
        If statut = "Honore" Then texte = texte & " (vu)"
    End If
    If Len(ws.Cells(ligne, col).Value) > 0 Then
        ws.Cells(ligne, col).Value = ws.Cells(ligne, col).Value & vbLf & texte
    Else
        ws.Cells(ligne, col).Value = texte
    End If
    ws.Cells(ligne, col).Interior.Color = CouleurStatut(statut)
    If blocage Then ws.Cells(ligne, col).Font.Italic = True
    If Not mCarte.Exists(cle) Then Set mCarte(cle) = New Collection
    mCarte(cle).Add rdv
    nbCreneaux = (Val(rdv("DureeMin")) + PasMinutes() - 1) \ PasMinutes()
    For k = 1 To nbCreneaux - 1
        If ligne + k > derniereLigne Then Exit For
        If Len(ws.Cells(ligne + k, col).Value) = 0 Then
            ws.Cells(ligne + k, col).Value = ChrW$(8627) & " " & nom
            ws.Cells(ligne + k, col).Interior.Color = CouleurStatut(statut)
            ws.Cells(ligne + k, col).Font.Color = RGB(120, 120, 120)
            If blocage Then ws.Cells(ligne + k, col).Font.Italic = True
            Dim cle2 As String
            cle2 = (ligne + k) & "|" & col
            If Not mCarte.Exists(cle2) Then Set mCarte(cle2) = New Collection
            mCarte(cle2).Add rdv
        End If
    Next k
End Sub

Private Function CouleurStatut(ByVal statut As String) As Long
    Select Case statut
        Case "Arrive": CouleurStatut = RGB(197, 224, 255)
        Case "Honore": CouleurStatut = RGB(217, 217, 217)
        Case "Absent": CouleurStatut = RGB(255, 199, 199)
        Case "Annule": CouleurStatut = RGB(230, 230, 230)
        Case modAgenda.STATUT_BLOQUE: CouleurStatut = RGB(200, 200, 200)
        Case Else:     CouleurStatut = RGB(255, 242, 180)
    End Select
End Function

' --- interaction (Worksheet_BeforeDoubleClick) --------------------------
Public Sub ClicCreneau(ByVal cible As Range)
    On Error GoTo Erreur
    Dim ligne As Long, col As Long, cle As String, d As Date, heure As String
    ligne = cible.Row: col = cible.Column
    Initialiser
    cle = ligne & "|" & col
    If mVueMois Then
        If mCarteJours Is Nothing Then Rendre
        If mCarteJours.Exists(cle) Then
            d = CDate(mCarteJours(cle))
            mVueMois = False
            mLundi = LundiDe(d)
            Rendre
        End If
        Exit Sub
    End If
    If mCarte Is Nothing Then Rendre
    If ligne < LIG_DEBUT Or col < COL_PREMIER_JOUR Or col > COL_PREMIER_JOUR + NbJours() - 1 Then Exit Sub
    heure = HeureDeLigne(ligne)
    If Len(heure) = 0 Then Exit Sub
    d = mLundi + (col - COL_PREMIER_JOUR)
    If mCarte.Exists(cle) Then
        ActionsCreneau mCarte(cle)
    Else
        PrendreRdv d, heure
    End If
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Public Function CreneauLibre(ByVal d As Date, ByVal heure As String, _
                             Optional ByVal dureeMin As String = "") As Boolean
    If Len(dureeMin) = 0 Then dureeMin = CStr(modConfig.ConfigNum("AGENDA", "DureeDefautMin", 15))
    CreneauLibre = (Len(modAgenda.ConflitRdv(Format$(d, "dd/mm/yyyy"), heure, dureeMin)) = 0)
End Function

Private Sub PrendreRdv(ByVal d As Date, ByVal heure As String)
    Dim f As ufRdvEdit
    Set f = New ufRdvEdit
    f.Prefixer Format$(d, "dd/mm/yyyy"), heure, mPatientPreselect
    f.Show vbModal
    Unload f
    Set mPatientPreselect = Nothing
    Rendre
End Sub

Private Sub ActionsCreneau(ByVal rdvs As Collection)
    Dim items As Collection, r As Object, it As Object, f As ufListe, nom As String
    Dim actions As Variant, a As Variant
    Set items = New Collection
    For Each r In rdvs
        If modAgenda.EstBlocage(r) Then
            nom = "INDISPONIBLE - " & r("Notes")
            actions = Array(Array("LEVER", "Lever l'indisponibilite (ce jour)"))
        Else
            If mPatients.Exists(r("PatientID")) Then
                nom = mPatients(r("PatientID"))("Prenom") & " " & mPatients(r("PatientID"))("Nom")
            Else
                nom = r("PatientID")
            End If
            actions = Array(Array("ARRIVE", "Marquer ARRIVE"), Array("ABSENT", "Marquer ABSENT"), _
                            Array("DEPLACER", "Deplacer / modifier"), _
                            Array("ANNULE", "Annuler le rendez-vous"), Array("FICHE", "Ouvrir la fiche patient"))
        End If
        For Each a In actions
            Set it = CreateObject("Scripting.Dictionary")
            it.CompareMode = 1
            it("ID") = r("ID") & "|" & a(0)
            it("Patient") = nom & " (" & r("Heure") & ")"
            it("Action") = a(1)
            items.Add it
        Next a
    Next r
    Set f = New ufListe
    f.Configurer "Rendez-vous : que faire ?", items, Array("Patient", "Action"), "170 pt;190 pt"
    f.Show vbModal
    If Not f.Annule Then
        Dim choix As String, rdvID As String, action As String
        choix = f.Resultat("ID")
        rdvID = Left$(choix, InStr(choix, "|") - 1)
        action = Mid$(choix, InStr(choix, "|") + 1)
        Unload f
        ExecuterAction rdvs, rdvID, action
    Else
        Unload f
    End If
End Sub

Private Sub ExecuterAction(ByVal rdvs As Collection, ByVal rdvID As String, ByVal action As String)
    Dim r As Object, x As Object, annee As Long, fp As ufPatientEdit, fr As ufRdvEdit
    For Each x In rdvs
        If x("ID") = rdvID Then Set r = x
    Next x
    If r Is Nothing Then Exit Sub
    annee = Year(DateDe(r("Date")))
    Select Case action
        Case "ARRIVE"
            modAgenda.MarquerStatut rdvID, "Arrive", annee
            If Len(modConfig.Config("ECG", "DossierGdt", "")) > 0 And mPatients.Exists(r("PatientID")) Then
                On Error Resume Next
                modGdt.EcrireGdtPatient mPatients(r("PatientID"))
                On Error GoTo 0
            End If
        Case "ABSENT": modAgenda.MarquerStatut rdvID, "Absent", annee
        Case "ANNULE"
            If MsgBox("Annuler ce rendez-vous ?", vbYesNo + vbQuestion, "Cabinet") = vbYes Then
                modAgenda.MarquerStatut rdvID, "Annule", annee
            End If
        Case "LEVER"
            If MsgBox("Lever l'indisponibilite du " & r("Date") & " (" & r("Notes") & ") ?", _
                      vbYesNo + vbQuestion, "Cabinet") = vbYes Then
                modAgenda.MarquerStatut rdvID, "Annule", annee
            End If
        Case "DEPLACER"
            Set fr = New ufRdvEdit
            If mPatients.Exists(r("PatientID")) Then
                fr.ChargerExistant r, mPatients(r("PatientID"))
            Else
                fr.ChargerExistant r, Nothing
            End If
            fr.Show vbModal
            Unload fr
        Case "FICHE"
            If mPatients.Exists(r("PatientID")) Then
                Set fp = New ufPatientEdit
                fp.ChargerExistant mPatients(r("PatientID"))
                fp.Show vbModal
                Unload fp
            End If
    End Select
    Rendre
End Sub
