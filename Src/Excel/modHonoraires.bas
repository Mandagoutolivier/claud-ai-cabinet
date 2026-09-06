Attribute VB_Name = "modHonoraires"
Option Explicit
' =====================================================================
' modHonoraires - Exploitation du journal des honoraires (poste secretaire),
' feuille "Honoraires", accessible par le bouton "Honoraires et
' encaissements" de l'Accueil. NE MODIFIE PAS le circuit d'enregistrement
' des seances : lecture du journal, plus trois ecritures ciblees
' (reglement via modActes.EnregistrerReglement, date de relance, date de
' remise d'un cheque).
'  - Impayes : seances non soldees, anciennete, relance
'  - Note d'honoraires / recu acquitte pour le patient (PDF + apercu)
'  - Remise de cheques : cheques recus non remis, bordereau
'  - Recettes du mois : encaissements SANS nom de patient (comptable)
'  - Tableau de bord : facture / encaisse / reste du, par mois et par acte
' =====================================================================

Private mVue As String                 ' IMPAYES / CHEQUES / RECETTES / BORD
Private mCarte As Object               ' ligne feuille -> SeanceID|annee
Private mSeances As Object             ' SeanceID -> dict agrege
Private mAnnee As Long

Private Const LIG_TITRE As Long = 2
Private Const LIG_ENTETE As Long = 3
Private Const LIG_DEBUT As Long = 4

' --- utilitaires --------------------------------------------------------
Private Function Nombre(ByVal v As Variant) As Double
    Nombre = Val(Replace(Trim$(CStr(v)), ",", "."))
End Function

Private Function Champ(ByVal d As Object, ByVal cle As String) As String
    If d.Exists(cle) Then Champ = Trim$(CStr(d(cle)))
End Function

Private Function DateDe(ByVal texte As String) As Date
    Dim p() As String
    p = Split(Trim$(Left$(texte, 10)), "/")
    If UBound(p) = 2 Then DateDe = DateSerial(Val(p(2)), Val(p(1)), Val(p(0))) Else DateDe = 0
End Function

Private Function EstCheque(ByVal mode As String) As Boolean
    EstCheque = (InStr(1, mode, "ch", vbTextCompare) = 1)   ' Cheque, CHQ, cheque...
End Function

Private Function Feuille() As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets("Honoraires")
    If Feuille Is Nothing Then
        Set Feuille = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Feuille.Name = "Honoraires"
    End If
End Function

Private Sub Nettoyer(ByVal ws As Worksheet)
    ws.Range(ws.Cells(LIG_TITRE, 1), ws.Cells(ws.Rows.Count, 12)).Clear
    ws.Range(ws.Cells(LIG_TITRE, 1), ws.Cells(ws.Rows.Count, 12)).RowHeight = 15
End Sub

Private Sub Titre(ByVal ws As Worksheet, ByVal t As String, ByVal aide As String)
    ws.Cells(LIG_TITRE, 1).Value = t
    ws.Cells(LIG_TITRE, 1).Font.Bold = True
    ws.Cells(LIG_TITRE, 1).Font.Size = 13
    ws.Cells(LIG_TITRE, 5).Value = aide
    ws.Cells(LIG_TITRE, 5).Font.Color = RGB(90, 90, 90)
End Sub

Private Sub Entetes(ByVal ws As Worksheet, ByVal noms As Variant, ByVal largeurs As Variant)
    Dim i As Long
    For i = LBound(noms) To UBound(noms)
        ws.Cells(LIG_ENTETE, i + 1).Value = noms(i)
        ws.Columns(i + 1).ColumnWidth = largeurs(i)
    Next i
    With ws.Range(ws.Cells(LIG_ENTETE, 1), ws.Cells(LIG_ENTETE, UBound(noms) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

' Lignes du journal d'une ou deux annees (l'annee demandee et la precedente :
' un impaye de decembre doit apparaitre en janvier)
Private Function LignesJournal(ByVal annee As Long, Optional ByVal avecPrecedente As Boolean = True) As Collection
    Dim res As Collection, l As Object, a As Long, tous As Collection
    Set res = New Collection
    For a = IIf(avecPrecedente, annee - 1, annee) To annee
        If modFichiers.FichierExiste(modConfig.FichierJournal(a)) Then
            modJournal.AssurerJournalAnnee a
            Set tous = Nothing
            On Error Resume Next
            Set tous = modBaseIO.LireTableX(modConfig.FichierJournal(a), "JOURNAL", "SeanceID")
            On Error GoTo 0
            If Not tous Is Nothing Then
                For Each l In tous
                    l("_Annee") = a
                    res.Add l
                Next l
            End If
        End If
    Next a
    Set LignesJournal = res
End Function

' Agregation par seance : Du, Regle, Reste, Actes, identite, dates
Private Function Seances(ByVal lignes As Collection) As Object
    Dim d As Object, l As Object, s As Object, id As String
    Set d = CreateObject("Scripting.Dictionary")
    For Each l In lignes
        id = Champ(l, "SeanceID")
        If Len(id) > 0 Then
            If Not d.Exists(id) Then
                Set s = CreateObject("Scripting.Dictionary")
                s("SeanceID") = id
                s("Annee") = l("_Annee")
                s("DateActe") = Champ(l, "DateActe")
                s("PatientID") = Champ(l, "PatientID")
                s("Nom") = Champ(l, "Nom")
                s("Prenom") = Champ(l, "Prenom")
                s("DDN") = Champ(l, "DDN")
                s("Actes") = ""
                s("Du") = 0#: s("Regle") = 0#
                s("Mode") = Champ(l, "ModePaiement")
                s("RelanceLe") = Champ(l, "RelanceLe")
                s("DateEncaissement") = Champ(l, "DateEncaissement")
                s("ChequeRemisLe") = Champ(l, "ChequeRemisLe")
                s("TiersPayant") = Champ(l, "TiersPayant")
                Set d(id) = s
            End If
            Set s = d(id)
            s("Du") = s("Du") + Nombre(Champ(l, "MontantDu"))
            s("Regle") = s("Regle") + Nombre(Champ(l, "MontantRegle"))
            s("Actes") = s("Actes") & IIf(Len(s("Actes")) > 0, " + ", "") & Champ(l, "CodeActe")
            If Len(Champ(l, "RelanceLe")) > 0 Then s("RelanceLe") = Champ(l, "RelanceLe")
        End If
    Next l
    Set Seances = d
End Function

' --- entree depuis l'Accueil --------------------------------------------
Public Sub Afficher()
    mAnnee = Year(Date)
    mVue = "IMPAYES"
    Rendre
End Sub

Public Sub HonorairesImpayes()
    mVue = "IMPAYES": Rendre
End Sub

Public Sub HonorairesCheques()
    mVue = "CHEQUES": Rendre
End Sub

Public Sub HonorairesRecettes()
    Dim s As String
    s = Trim$(InputBox("Mois des recettes (mm/aaaa) :", "Recettes du mois", Format$(Date, "mm/yyyy")))
    If Len(s) = 0 Then Exit Sub
    If Not s Like "##/####" Then MsgBox "Format attendu : mm/aaaa", vbExclamation, "Cabinet": Exit Sub
    RendreRecettes CLng(Left$(s, 2)), CLng(Right$(s, 4))
End Sub

Public Sub HonorairesTableauDeBord()
    Dim s As String
    s = Trim$(InputBox("Annee :", "Tableau de bord", CStr(Year(Date))))
    If Len(s) = 0 Then Exit Sub
    If Val(s) < 2000 Then Exit Sub
    mAnnee = CLng(s)
    mVue = "BORD": Rendre
End Sub

Public Sub HonorairesRetourAccueil()
    ThisWorkbook.Worksheets("Accueil").Activate
End Sub

Public Sub Rendre()
    On Error GoTo Erreur
    If mAnnee = 0 Then mAnnee = Year(Date)
    Application.ScreenUpdating = False
    Select Case mVue
        Case "CHEQUES": RendreCheques
        Case "BORD":    RendreTableauDeBord
        Case Else:      RendreImpayes
    End Select
    Application.ScreenUpdating = True
    Exit Sub
Erreur:
    Application.ScreenUpdating = True
    MsgBox "Affichage impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

' --- 1. impayes ---------------------------------------------------------
Private Sub RendreImpayes()
    Dim ws As Worksheet, k As Variant, s As Object, r As Long, reste As Double, total As Double
    Dim seuil As Long, age As Long, cles As Collection, tri As Object
    seuil = CLng(modConfig.ConfigNum("HONORAIRES", "RelanceApresJours", 30))
    Set ws = Feuille()
    Nettoyer ws
    Titre ws, "Impayes et encaissements - au " & Format$(Date, "dd/mm/yyyy"), _
          "Double-clic sur une ligne : enregistrer un reglement / relance / note d'honoraires / fiche"
    Entetes ws, Array("Date acte", "Patient", "Actes", "Du", "Regle", "Reste", "Anciennete", "Mode", "Relance le", "Tiers payant"), _
                Array(11, 30, 18, 9, 9, 9, 11, 16, 11, 10)
    Set mSeances = Seances(LignesJournal(mAnnee))
    Set mCarte = CreateObject("Scripting.Dictionary")
    ' tri par date d'acte croissante (les plus anciens en premier)
    Set cles = ClesTrieesParDate(mSeances)
    r = LIG_DEBUT
    For Each k In cles
        Set s = mSeances(k)
        reste = s("Du") - s("Regle")
        If reste > 0.001 Then
            age = Date - DateDe(s("DateActe"))
            ws.Cells(r, 1).Value = s("DateActe")
            ws.Cells(r, 2).Value = s("Nom") & " " & s("Prenom")
            ws.Cells(r, 3).Value = s("Actes")
            ws.Cells(r, 4).Value = s("Du")
            ws.Cells(r, 5).Value = s("Regle")
            ws.Cells(r, 6).Value = reste
            ws.Cells(r, 7).Value = age & " j"
            ws.Cells(r, 8).Value = s("Mode")
            ws.Cells(r, 9).Value = s("RelanceLe")
            ws.Cells(r, 10).Value = IIf(s("TiersPayant") = "O", "oui", "")
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 6)).NumberFormat = "# ##0.00"
            If age > seuil And Len(s("RelanceLe")) = 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = RGB(255, 199, 199)
            ElseIf s("TiersPayant") = "O" Or InStr(1, s("Mode"), "virement", vbTextCompare) > 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = RGB(255, 242, 180)
            End If
            mCarte(r) = s("SeanceID") & "|" & s("Annee")
            total = total + reste
            r = r + 1
        End If
    Next k
    ws.Cells(r + 1, 5).Value = "Total restant du :"
    ws.Cells(r + 1, 5).Font.Bold = True
    ws.Cells(r + 1, 6).Value = total
    ws.Cells(r + 1, 6).NumberFormat = "# ##0.00"
    ws.Cells(r + 1, 6).Font.Bold = True
    ws.Cells(r + 2, 1).Value = "Rouge : sans relance depuis plus de " & seuil & " jours. Jaune : virement attendu ou tiers payant."
    ws.Cells(r + 2, 1).Font.Color = RGB(90, 90, 90)
    If r = LIG_DEBUT Then ws.Cells(r, 2).Value = "(aucun impaye)"
    Finaliser ws
End Sub

Private Function ClesTrieesParDate(ByVal dict As Object) As Collection
    Dim col As Collection, k As Variant, i As Long, d As Date, inserted As Boolean
    Set col = New Collection
    For Each k In dict.Keys
        d = DateDe(dict(k)("DateActe"))
        inserted = False
        For i = 1 To col.Count
            If DateDe(dict(col(i))("DateActe")) > d Then
                col.Add k, , i: inserted = True: Exit For
            End If
        Next i
        If Not inserted Then col.Add k
    Next k
    Set ClesTrieesParDate = col
End Function

Private Sub Finaliser(ByVal ws As Worksheet)
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Cells(LIG_DEBUT, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

' --- 5. remise de cheques -----------------------------------------------
Private Sub RendreCheques()
    Dim ws As Worksheet, k As Variant, s As Object, r As Long, total As Long, montant As Double
    Set ws = Feuille()
    Nettoyer ws
    Titre ws, "Cheques recus non remis en banque", _
          "Double-clic sur une ligne : marquer remis en banque | bouton Bordereau : liste imprimable de tous les cheques a remettre"
    Entetes ws, Array("Date acte", "Patient", "Actes", "Montant", "Encaisse le", "Remis le"), _
                Array(11, 30, 18, 10, 12, 12)
    Set mSeances = Seances(LignesJournal(mAnnee))
    Set mCarte = CreateObject("Scripting.Dictionary")
    r = LIG_DEBUT
    For Each k In ClesTrieesParDate(mSeances)
        Set s = mSeances(k)
        If EstCheque(s("Mode")) And s("Regle") > 0.001 And Len(s("ChequeRemisLe")) = 0 Then
            ws.Cells(r, 1).Value = s("DateActe")
            ws.Cells(r, 2).Value = s("Nom") & " " & s("Prenom")
            ws.Cells(r, 3).Value = s("Actes")
            ws.Cells(r, 4).Value = s("Regle")
            ws.Cells(r, 4).NumberFormat = "# ##0.00"
            ws.Cells(r, 5).Value = s("DateEncaissement")
            mCarte(r) = s("SeanceID") & "|" & s("Annee")
            montant = montant + s("Regle")
            total = total + 1
            r = r + 1
        End If
    Next k
    ws.Cells(r + 1, 2).Value = total & " cheque(s), " & Format$(montant, "# ##0.00") & " EUR a remettre"
    ws.Cells(r + 1, 2).Font.Bold = True
    If r = LIG_DEBUT Then ws.Cells(r, 2).Value = "(aucun cheque en attente)"
    Finaliser ws
End Sub

' Bordereau de remise : tous les cheques non remis, PDF + apercu, puis
' marquage "remis le" apres confirmation.
Public Sub HonorairesBordereauCheques()
    On Error GoTo Erreur
    Dim ws As Worksheet, k As Variant, s As Object, r As Long, n As Long, montant As Double, pdf As String
    Dim aMarquer As Collection
    If mAnnee = 0 Then mAnnee = Year(Date)
    Set mSeances = Seances(LignesJournal(mAnnee))
    Set aMarquer = New Collection
    Set ws = FeuilleTemp("Bordereau")
    ws.Cells(1, 1).Value = modConfig.Config("GENERAL", "NomCabinet", "Cabinet") & " - Bordereau de remise de cheques du " & Format$(Date, "dd/mm/yyyy")
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
    ws.Cells(3, 1).Value = "N°": ws.Cells(3, 2).Value = "Date acte": ws.Cells(3, 3).Value = "Tireur"
    ws.Cells(3, 4).Value = "Montant": ws.Cells(3, 5).Value = "Seance"
    ws.Range("A3:E3").Font.Bold = True
    r = 4
    For Each k In ClesTrieesParDate(mSeances)
        Set s = mSeances(k)
        If EstCheque(s("Mode")) And s("Regle") > 0.001 And Len(s("ChequeRemisLe")) = 0 Then
            n = n + 1
            ws.Cells(r, 1).Value = n
            ws.Cells(r, 2).Value = s("DateActe")
            ws.Cells(r, 3).Value = s("Nom") & " " & s("Prenom")
            ws.Cells(r, 4).Value = s("Regle"): ws.Cells(r, 4).NumberFormat = "# ##0.00"
            ws.Cells(r, 5).Value = s("SeanceID")
            montant = montant + s("Regle")
            aMarquer.Add s
            r = r + 1
        End If
    Next k
    If n = 0 Then MsgBox "Aucun cheque a remettre.", vbInformation, "Cabinet": Exit Sub
    ws.Cells(r + 1, 3).Value = "Total : " & n & " cheque(s)"
    ws.Cells(r + 1, 4).Value = montant: ws.Cells(r + 1, 4).NumberFormat = "# ##0.00"
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 1, 4)).Font.Bold = True
    ws.Columns(1).ColumnWidth = 5: ws.Columns(2).ColumnWidth = 12: ws.Columns(3).ColumnWidth = 32
    ws.Columns(4).ColumnWidth = 12: ws.Columns(5).ColumnWidth = 24
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r + 1, 5)).Address
    pdf = DossierSortie("Remises") & "\Remise_cheques_" & Format$(Now, "yyyy-mm-dd_hhnn") & ".pdf"
    On Error Resume Next
    ws.ExportAsFixedFormat 0, pdf
    On Error GoTo Erreur
    ws.Activate
    ws.PrintPreview
    If MsgBox("Marquer ces " & n & " cheque(s) comme REMIS EN BANQUE aujourd'hui ?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Cabinet") = vbYes Then
        Dim v As Object
        For Each s In aMarquer
            Set v = CreateObject("Scripting.Dictionary")
            v("ChequeRemisLe") = Format$(Date, "dd/mm/yyyy")
            modJournal.MettreAJourSeance s("SeanceID"), v, s("Annee")
        Next s
        MsgBox n & " cheque(s) marques remis. Bordereau : " & pdf, vbInformation, "Cabinet"
    End If
    mVue = "CHEQUES": Rendre
    Exit Sub
Erreur:
    MsgBox "Bordereau impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

' --- 3. recettes du mois (sans nom de patient) ----------------------------
Private Sub RendreRecettes(ByVal mois As Long, ByVal annee As Long)
    On Error GoTo Erreur
    Dim ws As Worksheet, l As Object, r As Long, d As Date, montant As Double
    Dim parMode As Object, parJour As Object, k As Variant, total As Double, pdf As String, xlsx As String
    Dim lignes As Collection, tri As Collection, i As Long, inserted As Boolean
    Set parMode = CreateObject("Scripting.Dictionary")
    Set parJour = CreateObject("Scripting.Dictionary")
    ' lignes encaissees dans le mois, triees par date d'encaissement
    Set tri = New Collection
    For Each l In LignesJournal(annee, False)
        montant = Nombre(Champ(l, "MontantRegle"))
        d = DateDe(Champ(l, "DateEncaissement"))
        If montant > 0.001 And d <> 0 Then
            If Month(d) = mois And Year(d) = annee Then
                inserted = False
                For i = 1 To tri.Count
                    If DateDe(Champ(tri(i), "DateEncaissement")) > d Then tri.Add l, , i: inserted = True: Exit For
                Next i
                If Not inserted Then tri.Add l
            End If
        End If
    Next l
    Set ws = FeuilleTemp("Recettes")
    ws.Cells(1, 1).Value = modConfig.Config("GENERAL", "NomCabinet", "Cabinet") & " - Recettes de " & Format$(DateSerial(annee, mois, 1), "mmmm yyyy")
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
    ws.Cells(2, 1).Value = "Encaissements effectivement recus (honoraires), sans identite de patient. Edite le " & Format$(Now, "dd/mm/yyyy hh:nn")
    ws.Cells(2, 1).Font.Color = RGB(90, 90, 90)
    ws.Cells(4, 1).Value = "Date encaissement": ws.Cells(4, 2).Value = "Mode": ws.Cells(4, 3).Value = "Acte"
    ws.Cells(4, 4).Value = "Montant": ws.Cells(4, 5).Value = "Reference seance": ws.Cells(4, 6).Value = "Date acte"
    ws.Range("A4:F4").Font.Bold = True
    ws.Columns(1).NumberFormat = "@": ws.Columns(6).NumberFormat = "@"
    r = 5
    For Each l In tri
        montant = Nombre(Champ(l, "MontantRegle"))
        ws.Cells(r, 1).Value = Champ(l, "DateEncaissement")
        ws.Cells(r, 2).Value = Champ(l, "ModePaiement")
        ws.Cells(r, 3).Value = Champ(l, "CodeOfficiel")
        ws.Cells(r, 4).Value = montant: ws.Cells(r, 4).NumberFormat = "# ##0.00"
        ws.Cells(r, 5).Value = Champ(l, "SeanceID")
        ws.Cells(r, 6).Value = Champ(l, "DateActe")
        parMode(Champ(l, "ModePaiement")) = parMode(Champ(l, "ModePaiement")) + montant
        parJour(Champ(l, "DateEncaissement")) = parJour(Champ(l, "DateEncaissement")) + montant
        total = total + montant
        r = r + 1
    Next l
    r = r + 1
    ws.Cells(r, 3).Value = "TOTAL DU MOIS": ws.Cells(r, 4).Value = total
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Font.Bold = True: ws.Cells(r, 4).NumberFormat = "# ##0.00"
    r = r + 2
    ws.Cells(r, 1).Value = "Par mode de paiement": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    For Each k In parMode.Keys
        ws.Cells(r, 2).Value = k: ws.Cells(r, 4).Value = parMode(k): ws.Cells(r, 4).NumberFormat = "# ##0.00": r = r + 1
    Next k
    r = r + 1
    ws.Cells(r, 1).Value = "Par jour": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    For Each k In parJour.Keys
        ws.Cells(r, 1).Value = k: ws.Cells(r, 4).Value = parJour(k): ws.Cells(r, 4).NumberFormat = "# ##0.00": r = r + 1
    Next k
    ws.Columns(1).ColumnWidth = 17: ws.Columns(2).ColumnWidth = 16: ws.Columns(3).ColumnWidth = 10
    ws.Columns(4).ColumnWidth = 11: ws.Columns(5).ColumnWidth = 24: ws.Columns(6).ColumnWidth = 11
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 6)).Address
    ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesTall = False
    ' exports : PDF et classeur Excel (pour le comptable)
    pdf = DossierSortie("Recettes") & "\Recettes_" & Format$(DateSerial(annee, mois, 1), "yyyy-mm") & ".pdf"
    xlsx = DossierSortie("Recettes") & "\Recettes_" & Format$(DateSerial(annee, mois, 1), "yyyy-mm") & ".xlsx"
    On Error Resume Next
    ws.ExportAsFixedFormat 0, pdf
    Dim wb As Workbook
    Application.DisplayAlerts = False
    ws.Copy
    Set wb = ActiveWorkbook
    wb.SaveAs xlsx, 51
    wb.Close False
    Application.DisplayAlerts = True
    On Error GoTo Erreur
    ws.Activate
    MsgBox "Recettes de " & Format$(DateSerial(annee, mois, 1), "mmmm yyyy") & " : " & Format$(total, "# ##0.00") & " EUR" & vbCrLf & _
           "PDF   : " & pdf & vbCrLf & "Excel : " & xlsx, vbInformation, "Cabinet"
    Exit Sub
Erreur:
    Application.DisplayAlerts = True
    MsgBox "Recettes impossibles : " & Err.Description, vbExclamation, "Cabinet"
End Sub

' --- 4. tableau de bord --------------------------------------------------
Private Sub RendreTableauDeBord()
    Dim ws As Worksheet, l As Object, m As Long, r As Long, d As Date
    Dim facture(1 To 12) As Double, encaisse(1 To 12) As Double, resteMois(1 To 12) As Double
    Dim parActe As Object, nbActe As Object, k As Variant, montant As Double, du As Double, regle As Double
    Dim totF As Double, totE As Double, totR As Double
    Set parActe = CreateObject("Scripting.Dictionary")
    Set nbActe = CreateObject("Scripting.Dictionary")
    For Each l In LignesJournal(mAnnee, False)
        du = Nombre(Champ(l, "MontantDu")): regle = Nombre(Champ(l, "MontantRegle"))
        d = DateDe(Champ(l, "DateActe"))
        If d <> 0 Then
            If Year(d) = mAnnee Then
                m = Month(d)
                facture(m) = facture(m) + du
                resteMois(m) = resteMois(m) + (du - regle)
                parActe(Champ(l, "CodeActe")) = parActe(Champ(l, "CodeActe")) + du
                nbActe(Champ(l, "CodeActe")) = nbActe(Champ(l, "CodeActe")) + 1
            End If
        End If
        d = DateDe(Champ(l, "DateEncaissement"))
        If d <> 0 And regle > 0.001 Then
            If Year(d) = mAnnee Then encaisse(Month(d)) = encaisse(Month(d)) + regle
        End If
    Next l
    Set ws = Feuille()
    Nettoyer ws
    Titre ws, "Tableau de bord " & mAnnee, "Facture = honoraires des actes du mois ; Encaisse = sommes recues dans le mois ; Reste = impayes des actes du mois"
    Entetes ws, Array("Mois", "Facture", "Encaisse", "Reste du", "", "Acte", "Nombre", "Facture"), _
                Array(12, 12, 12, 12, 3, 12, 9, 12)
    Set mCarte = CreateObject("Scripting.Dictionary")
    For m = 1 To 12
        r = LIG_DEBUT + m - 1
        ws.Cells(r, 1).Value = Format$(DateSerial(mAnnee, m, 1), "mmmm")
        ws.Cells(r, 2).Value = facture(m): ws.Cells(r, 3).Value = encaisse(m): ws.Cells(r, 4).Value = resteMois(m)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).NumberFormat = "# ##0.00"
        If resteMois(m) > 0.001 Then ws.Cells(r, 4).Font.Color = RGB(192, 0, 0)
        totF = totF + facture(m): totE = totE + encaisse(m): totR = totR + resteMois(m)
    Next m
    r = LIG_DEBUT + 12
    ws.Cells(r, 1).Value = "TOTAL": ws.Cells(r, 2).Value = totF: ws.Cells(r, 3).Value = totE: ws.Cells(r, 4).Value = totR
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).NumberFormat = "# ##0.00"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
    r = LIG_DEBUT
    For Each k In parActe.Keys
        ws.Cells(r, 6).Value = k
        ws.Cells(r, 7).Value = nbActe(k)
        ws.Cells(r, 8).Value = parActe(k): ws.Cells(r, 8).NumberFormat = "# ##0.00"
        r = r + 1
    Next k
    Finaliser ws
End Sub

' --- interaction : double-clic sur une ligne -----------------------------
Public Sub ClicLigne(ByVal cible As Range)
    On Error GoTo Erreur
    Dim cle As String, seanceID As String, annee As Long, s As Object
    If mCarte Is Nothing Then Exit Sub
    If Not mCarte.Exists(cible.Row) Then Exit Sub
    cle = mCarte(cible.Row)
    seanceID = Left$(cle, InStr(cle, "|") - 1)
    annee = CLng(Mid$(cle, InStr(cle, "|") + 1))
    Set s = mSeances(seanceID)
    If mVue = "CHEQUES" Then
        If MsgBox("Marquer le cheque de " & s("Nom") & " " & s("Prenom") & " (" & Format$(s("Regle"), "0.00") & _
                  " EUR) comme remis en banque aujourd'hui ?", vbYesNo + vbQuestion, "Cabinet") = vbYes Then
            Dim v As Object
            Set v = CreateObject("Scripting.Dictionary")
            v("ChequeRemisLe") = Format$(Date, "dd/mm/yyyy")
            modJournal.MettreAJourSeance seanceID, v, annee
            Rendre
        End If
        Exit Sub
    End If
    ActionsSeance s
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub ActionsSeance(ByVal s As Object)
    Dim items As Collection, it As Object, f As ufListe, a As Variant, reste As Double
    reste = s("Du") - s("Regle")
    Set items = New Collection
    For Each a In Array(Array("REGLEMENT", "Enregistrer un reglement (reste " & Format$(reste, "0.00") & " EUR)"), _
                        Array("RELANCE", "Marquer relance aujourd'hui"), _
                        Array("NOTE", "Note d'honoraires / recu pour le patient"), _
                        Array("FICHE", "Ouvrir la fiche patient"))
        Set it = CreateObject("Scripting.Dictionary")
        it.CompareMode = 1
        it("ID") = a(0)
        it("Patient") = s("Nom") & " " & s("Prenom") & " - " & s("DateActe")
        it("Action") = a(1)
        items.Add it
    Next a
    Set f = New ufListe
    f.Configurer "Seance : que faire ?", items, Array("Patient", "Action"), "170 pt;230 pt"
    f.Show vbModal
    If f.Annule Then Unload f: Exit Sub
    a = f.Resultat("ID")
    Unload f
    Select Case a
        Case "REGLEMENT": SaisirReglement s
        Case "RELANCE"
            Dim v As Object
            Set v = CreateObject("Scripting.Dictionary")
            v("RelanceLe") = Format$(Date, "dd/mm/yyyy")
            modJournal.MettreAJourSeance s("SeanceID"), v, s("Annee")
            Rendre
        Case "NOTE": NoteHonoraires s("SeanceID"), s("Annee")
        Case "FICHE": OuvrirFiche s("PatientID")
    End Select
End Sub

Private Sub SaisirReglement(ByVal s As Object)
    On Error GoTo Erreur
    Dim reste As Double, txt As String, montant As Double, mode As String, solde As Double, i As Long
    Dim modes As Variant, choix As String
    reste = s("Du") - s("Regle")
    txt = Trim$(InputBox("Montant recu (EUR) pour " & s("Nom") & " " & s("Prenom") & vbCrLf & _
                         "Reste du : " & Format$(reste, "0.00") & " EUR", "Reglement", Format$(reste, "0.00")))
    If Len(txt) = 0 Then Exit Sub
    montant = Nombre(txt)
    If montant <= 0 Then MsgBox "Montant invalide.", vbExclamation, "Cabinet": Exit Sub
    modes = Array("CB", "Cheque", "Especes", "Virement")
    choix = Trim$(InputBox("Mode de paiement :" & vbCrLf & "1 = CB   2 = Cheque   3 = Especes   4 = Virement", "Reglement", "1"))
    If Len(choix) = 0 Then Exit Sub
    i = Val(choix)
    If i < 1 Or i > 4 Then mode = choix Else mode = modes(i - 1)
    solde = modActes.EnregistrerReglement(s("SeanceID"), montant, mode, s("Annee"))
    MsgBox "Reglement enregistre : " & Format$(montant, "0.00") & " EUR (" & mode & ")." & vbCrLf & _
           IIf(solde > 0.001, "Reste du : " & Format$(solde, "0.00") & " EUR", "Seance soldee."), vbInformation, "Cabinet"
    Rendre
    Exit Sub
Erreur:
    MsgBox "Reglement non enregistre : " & Err.Description, vbCritical, "Cabinet"
End Sub

Private Sub OuvrirFiche(ByVal patientID As String)
    Dim p As Object, fp As ufPatientEdit
    On Error Resume Next
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If p("ID") = patientID Then
            Set fp = New ufPatientEdit
            fp.ChargerExistant p
            fp.Show vbModal
            Unload fp
            Exit Sub
        End If
    Next p
End Sub

' --- 2. note d'honoraires / recu -----------------------------------------
' Document remis au patient (mutuelle, justificatif) : identite, actes,
' montants, mention "acquittee" si soldee. Numero = SeanceID. PDF dans
' Echange\Notes + apercu avant impression. AUCUNE ecriture au journal.
Public Sub NoteHonoraires(ByVal seanceID As String, ByVal annee As Long)
    On Error GoTo Erreur
    Dim ws As Worksheet, lignes As Collection, l As Object, r As Long, du As Double, regle As Double
    Dim pat As Object, p As Object, libelles As Object, a As Object, pdf As String, nom As String
    Set lignes = modJournal.LignesSeance(seanceID, annee)
    If lignes.Count = 0 Then MsgBox "Seance introuvable : " & seanceID, vbExclamation, "Cabinet": Exit Sub
    ' libelles des actes
    Set libelles = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    For Each a In modActes.Nomenclature()
        libelles(a("Code")) = a("LibelleCourt")
        If Len(a("CodeAssocie")) > 0 And Not libelles.Exists(a("CodeAssocie")) Then libelles(a("CodeAssocie")) = "Acte associe"
    Next a
    For Each p In modBaseIO.LireTableX(modConfig.FichierPatients(), "PATIENTS")
        If p("ID") = lignes(1)("PatientID") Then Set pat = p: Exit For
    Next p
    On Error GoTo Erreur

    Set ws = FeuilleTemp("Note")
    ws.Cells(1, 1).Value = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom")
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
    ws.Cells(2, 1).Value = modConfig.Config("MEDECIN", "Specialite")
    ws.Cells(3, 1).Value = modConfig.Config("MEDECIN", "AdresseLigne1") & " - " & modConfig.Config("MEDECIN", "AdresseLigne2")
    ws.Cells(4, 1).Value = "Tel : " & modConfig.Config("MEDECIN", "Telephone") & "   RPPS : " & modConfig.Config("MEDECIN", "RPPS") & _
                           "   N° AM : " & modConfig.Config("MEDECIN", "NumeroAM")
    ws.Cells(6, 1).Value = "NOTE D'HONORAIRES n° " & seanceID
    ws.Cells(6, 1).Font.Bold = True: ws.Cells(6, 1).Font.Size = 14
    ws.Cells(7, 1).Value = "Editee le " & Format$(Date, "dd/mm/yyyy")
    nom = Champ(lignes(1), "Nom") & " " & Champ(lignes(1), "Prenom")
    ws.Cells(9, 1).Value = "Patient : " & nom & "   ne(e) le " & Champ(lignes(1), "DDN")
    ws.Cells(9, 1).Font.Bold = True
    If Not pat Is Nothing Then
        ws.Cells(10, 1).Value = Trim$(Champ(pat, "Adresse1") & " " & Champ(pat, "Adresse2") & "  " & Champ(pat, "CP") & " " & Champ(pat, "Ville"))
        If Len(Champ(pat, "NIR")) > 0 Then ws.Cells(11, 1).Value = "N° securite sociale : " & Champ(pat, "NIR")
    End If
    ws.Cells(13, 1).Value = "Date": ws.Cells(13, 2).Value = "Code": ws.Cells(13, 3).Value = "Acte"
    ws.Cells(13, 4).Value = "Honoraires": ws.Cells(13, 5).Value = "Regle"
    ws.Range("A13:E13").Font.Bold = True
    ws.Range("A13:E13").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Columns(1).NumberFormat = "@"
    r = 14
    For Each l In lignes
        ws.Cells(r, 1).Value = Champ(l, "DateActe")
        ws.Cells(r, 2).Value = Champ(l, "CodeOfficiel")
        ws.Cells(r, 3).Value = IIf(libelles.Exists(Champ(l, "CodeActe")), libelles(Champ(l, "CodeActe")), Champ(l, "CodeActe"))
        ws.Cells(r, 4).Value = Nombre(Champ(l, "MontantDu")): ws.Cells(r, 4).NumberFormat = "# ##0.00"
        ws.Cells(r, 5).Value = Nombre(Champ(l, "MontantRegle")): ws.Cells(r, 5).NumberFormat = "# ##0.00"
        du = du + Nombre(Champ(l, "MontantDu")): regle = regle + Nombre(Champ(l, "MontantRegle"))
        r = r + 1
    Next l
    r = r + 1
    ws.Cells(r, 3).Value = "TOTAL": ws.Cells(r, 4).Value = du: ws.Cells(r, 5).Value = regle
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Font.Bold = True
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).NumberFormat = "# ##0.00"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
    r = r + 2
    If du - regle <= 0.001 Then
        ws.Cells(r, 1).Value = "Acquittee le " & Champ(lignes(1), "DateEncaissement") & " - mode de paiement : " & Champ(lignes(1), "ModePaiement")
    Else
        ws.Cells(r, 1).Value = "Reste du : " & Format$(du - regle, "# ##0.00") & " EUR"
    End If
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r + 2, 1).Value = "Document remis au patient a sa demande (mutuelle, justificatif). Il ne remplace pas la feuille de soins."
    ws.Cells(r + 2, 1).Font.Color = RGB(90, 90, 90)
    ws.Cells(r + 5, 4).Value = modConfig.Config("MEDECIN", "Signature", "")
    ws.Columns(1).ColumnWidth = 12: ws.Columns(2).ColumnWidth = 10: ws.Columns(3).ColumnWidth = 40
    ws.Columns(4).ColumnWidth = 12: ws.Columns(5).ColumnWidth = 12
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r + 6, 5)).Address
    ws.PageSetup.Orientation = xlPortrait
    pdf = DossierSortie("Notes") & "\Note_" & seanceID & ".pdf"
    On Error Resume Next
    ws.ExportAsFixedFormat 0, pdf
    On Error GoTo Erreur
    ws.Activate
    ws.PrintPreview
    If Len(Dir$(pdf)) > 0 Then MsgBox "Note d'honoraires : " & pdf, vbInformation, "Cabinet"
    Feuille().Activate
    Exit Sub
Erreur:
    MsgBox "Note d'honoraires impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

' --- feuilles de travail et dossiers de sortie ---------------------------
Private Function FeuilleTemp(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set FeuilleTemp = ThisWorkbook.Worksheets(nom)
    If FeuilleTemp Is Nothing Then
        Set FeuilleTemp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        FeuilleTemp.Name = nom
    End If
    FeuilleTemp.Cells.Clear
    FeuilleTemp.PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
    FeuilleTemp.PageSetup.RightMargin = Application.CentimetersToPoints(1.5)
End Function

Private Function DossierSortie(ByVal sousDossier As String) As String
    DossierSortie = modConfig.Chemin("Actes") & "\" & sousDossier
    modFichiers.EnsureDossier DossierSortie
End Function
