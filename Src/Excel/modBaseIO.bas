Attribute VB_Name = "modBaseIO"
Option Explicit
' =====================================================================
' modBaseIO - Ecritures dans les bases (poste secretaire UNIQUEMENT).
' Toute ecriture est une TRANSACTION COURTE : verrou cooperatif ->
' sauvegarde (1re ecriture du jour) -> ouvrir -> ecrire -> fermer ->
' relacher. Les fichiers ne restent JAMAIS ouverts.
' =====================================================================

Private mJourSauvegarde As String   ' "yyyymmdd|fichier" deja sauvegardes

' --- lecture (copie locale, fenetre cachee) ---------------------------

Public Function LireTableX(ByVal fichier As String, ByVal feuille As String, _
                           Optional ByVal colonneNonVide As String = "ID") As Collection
    Dim wb As Workbook, col As Collection
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(modFichiers.CopieLocale(fichier), ReadOnly:=True, AddToMru:=False)
    wb.Windows(1).Visible = False
    Set col = LireFeuilleX(wb, feuille, colonneNonVide)
    wb.Close False
    Application.ScreenUpdating = True
    Set LireTableX = col
End Function

Private Function LireFeuilleX(ByVal wb As Workbook, ByVal feuille As String, _
                              ByVal colonneNonVide As String) As Collection
    Dim ws As Worksheet, donnees As Variant, col As Collection, dict As Object
    Dim r As Long, c As Long
    Set ws = wb.Worksheets(feuille)
    Set col = New Collection
    donnees = ws.UsedRange.Value
    If IsEmpty(donnees) Or Not IsArray(donnees) Then Set LireFeuilleX = col: Exit Function
    For r = 2 To UBound(donnees, 1)
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1
        For c = 1 To UBound(donnees, 2)
            dict(Nz(donnees(1, c))) = Nz(donnees(r, c))
        Next c
        If Len(colonneNonVide) = 0 Then
            col.Add dict
        ElseIf dict.Exists(colonneNonVide) Then
            If Len(dict(colonneNonVide)) > 0 Then col.Add dict
        End If
    Next r
    Set LireFeuilleX = col
End Function

Private Function Nz(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        Nz = ""
    ElseIf VarType(v) = vbDate Then
        Nz = Format$(v, "dd/mm/yyyy")
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

' --- transactions d'ecriture -----------------------------------------

' Ajoute une ligne. Si prefixeID est fourni ("P", "C", "R"...), genere
' l'identifiant (colonne ID) et le renvoie.
Public Function AjouterLigne(ByVal fichier As String, ByVal feuille As String, _
                             ByVal valeurs As Object, Optional ByVal prefixeID As String = "") As String
    Dim wb As Workbook, ws As Worksheet, ligne As Long, id As String
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(feuille)
    ligne = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1      ' xlUp
    If ligne < 2 Then ligne = 2
    If Len(prefixeID) > 0 Then
        id = ProchainID(ws, prefixeID)
        valeurs("ID") = id
    End If
    EcrireValeurs ws, ligne, valeurs
    On Error GoTo 0
    FermerTransaction fichier, wb, True     ' leve si l'enregistrement echoue
    AjouterLigne = id
    Exit Function
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Function

' Modifie la premiere ligne ou colonneCle = valeurCle.
Public Sub ModifierLigne(ByVal fichier As String, ByVal feuille As String, _
                         ByVal colonneCle As String, ByVal valeurCle As String, _
                         ByVal valeurs As Object)
    Dim wb As Workbook, ws As Worksheet, ligne As Long
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(feuille)
    ligne = TrouverLigne(ws, colonneCle, valeurCle)
    If ligne = 0 Then Err.Raise vbObjectError + 600, "modBaseIO", _
        "Introuvable : " & colonneCle & "=" & valeurCle & " dans " & feuille
    EcrireValeurs ws, ligne, valeurs
    On Error GoTo 0
    FermerTransaction fichier, wb, True     ' leve si l'enregistrement echoue
    Exit Sub
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Sub

' Ajoute plusieurs lignes d'un coup (journal comptable)
Public Sub AjouterLignes(ByVal fichier As String, ByVal feuille As String, ByVal lignes As Collection)
    Dim wb As Workbook, ws As Worksheet, ligne As Long, valeurs As Object
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(feuille)
    ligne = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
    If ligne < 2 Then ligne = 2
    For Each valeurs In lignes
        EcrireValeurs ws, ligne, valeurs
        ligne = ligne + 1
    Next valeurs
    On Error GoTo 0
    FermerTransaction fichier, wb, True     ' leve si l'enregistrement echoue
    Exit Sub
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Sub

' Modifie TOUTES les lignes ou colonneCle = valeurCle (ex : toutes les
' lignes du journal d'une meme seance). Renvoie le nombre de lignes modifiees.
Public Function ModifierLignes(ByVal fichier As String, ByVal feuille As String, _
                               ByVal colonneCle As String, ByVal valeurCle As String, _
                               ByVal valeurs As Object) As Long
    Dim wb As Workbook, ws As Worksheet, lignes As Collection, l As Variant
    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(feuille)
    Set lignes = TrouverLignes(ws, colonneCle, valeurCle)
    For Each l In lignes
        EcrireValeurs ws, CLng(l), valeurs
    Next l
    ModifierLignes = lignes.Count
    On Error GoTo 0
    FermerTransaction fichier, wb, True
    Exit Function
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Function

' Ajoute a la premiere ligne les entetes manquants (migration d'un
' classeur cree par une version anterieure). Ne touche a aucune donnee.
Public Sub AssurerColonnes(ByVal fichier As String, ByVal feuille As String, ByVal entetes As Variant)
    Dim wb As Workbook, ws As Worksheet, presents As Object, c As Long, i As Long
    Dim derniere As Long, manque As Boolean
    If Not modFichiers.FichierExiste(fichier) Then Exit Sub
    ' lecture prealable sans verrou : n'ouvrir en ecriture que si necessaire
    Set presents = CreateObject("Scripting.Dictionary")
    presents.CompareMode = 1
    On Error Resume Next
    Set wb = Workbooks.Open(modFichiers.CopieLocale(fichier), ReadOnly:=True, AddToMru:=False)
    If Err.Number <> 0 Then Exit Sub
    wb.Windows(1).Visible = False
    On Error GoTo 0
    Set ws = wb.Worksheets(feuille)
    derniere = ws.Cells(1, ws.Columns.Count).End(-4159).Column
    For c = 1 To derniere
        presents(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    wb.Close False
    For i = LBound(entetes) To UBound(entetes)
        If Not presents.Exists(CStr(entetes(i))) Then manque = True
    Next i
    If Not manque Then Exit Sub

    OuvrirTransaction fichier, wb
    On Error GoTo Echec
    Set ws = wb.Worksheets(feuille)
    derniere = ws.Cells(1, ws.Columns.Count).End(-4159).Column
    For i = LBound(entetes) To UBound(entetes)
        If Not presents.Exists(CStr(entetes(i))) Then
            derniere = derniere + 1
            ws.Cells(1, derniere).Value = CStr(entetes(i))
            ws.Cells(1, derniere).Font.Bold = True
            presents(CStr(entetes(i))) = derniere
        End If
    Next i
    On Error GoTo 0
    FermerTransaction fichier, wb, True
    modLog.LogInfo "Colonnes completees dans " & fichier & " (" & feuille & ")"
    Exit Sub
Echec:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    FermerTransaction fichier, wb, False
    Err.Raise numErr, "modBaseIO", descErr
End Sub

' --- plomberie --------------------------------------------------------

Private Sub OuvrirTransaction(ByVal fichier As String, ByRef wb As Workbook)
    ' NB : ne jamais nommer une variable comme une fonction du module
    ' (VBA est insensible a la casse : elle masquerait la fonction)
    Dim verrou As String
    modLog.LogInfo "transaction: verrou " & fichier
    verrou = NomVerrou(fichier)
    If Not modFichiers.AcquerirVerrou(verrou, 8000) Then
        Err.Raise vbObjectError + 601, "modBaseIO", _
            "La base est occupee par un autre poste : reessayez dans quelques secondes."
    End If
    modLog.LogInfo "transaction: sauvegarde"
    SauvegardeQuotidienne fichier
    Application.ScreenUpdating = False
    On Error GoTo EchecOuverture
    modLog.LogInfo "transaction: ouverture"
    Set wb = Workbooks.Open(fichier, ReadOnly:=False, AddToMru:=False)
    modLog.LogInfo "transaction: ouverte, masquage fenetre"
    wb.Windows(1).Visible = False
    On Error GoTo 0
    modLog.LogInfo "transaction: prete"
    Exit Sub
EchecOuverture:
    Dim descErr As String
    descErr = Err.Description
    Application.ScreenUpdating = True
    modFichiers.RelacherVerrou verrou
    Err.Raise vbObjectError + 602, "modBaseIO", _
        "Impossible d'ouvrir " & fichier & " en ecriture (fichier ouvert sur un autre poste ?) : " & descErr
End Sub

' Ferme la transaction. enregistrer=True : l'enregistrement DOIT reussir,
' sinon l'erreur est propagee a l'appelant (aucun "succes" silencieux :
' la tache doit rester en attente et le message parvenir a la secretaire).
Private Sub FermerTransaction(ByVal fichier As String, ByVal wb As Workbook, ByVal enregistrer As Boolean)
    Dim numErr As Long, descErr As String, ferme As Boolean
    If wb Is Nothing Then
        Application.ScreenUpdating = True
        modFichiers.RelacherVerrou NomVerrou(fichier)
        Exit Sub
    End If

    If enregistrer Then
        ' la fenetre doit etre visible AVANT l'enregistrement : Excel memorise
        ' l'etat masque dans le fichier, qui s'ouvrirait ensuite "sans fenetre"
        On Error Resume Next
        wb.Windows(1).Visible = True
        Err.Clear
        wb.Save                                  ' <- l'echec est capture ici
        numErr = Err.Number: descErr = Err.Description
        On Error GoTo 0
        If numErr = 0 Then
            On Error Resume Next
            wb.Close SaveChanges:=False          ' deja enregistre
            ferme = (Err.Number = 0)
            On Error GoTo 0
        Else
            ' echec d'enregistrement : fermer SANS enregistrer et signaler
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
        End If
    Else
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If

    Application.ScreenUpdating = True
    modFichiers.RelacherVerrou NomVerrou(fichier)

    If enregistrer And numErr <> 0 Then
        modLog.LogErreur "Enregistrement refuse : " & fichier & " : " & descErr
        Err.Raise vbObjectError + 603, "modBaseIO", _
            "L'enregistrement de " & fichier & " a ECHOUE : " & descErr & vbCrLf & _
            "Rien n'a ete enregistre. Verifiez l'acces au dossier partage puis recommencez."
    End If
End Sub

Private Function NomVerrou(ByVal fichier As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    NomVerrou = fso.GetBaseName(fichier)
End Function

Private Sub SauvegardeQuotidienne(ByVal fichier As String)
    Dim cle As String
    cle = Format$(Date, "yyyymmdd") & "|" & fichier
    If InStr(mJourSauvegarde, cle) = 0 Then
        modFichiers.SauvegardeHorodatee fichier
        mJourSauvegarde = mJourSauvegarde & ";" & cle
    End If
End Sub

Private Sub EcrireValeurs(ByVal ws As Worksheet, ByVal ligne As Long, ByVal valeurs As Object)
    Dim entetes As Object, c As Long, cle As Variant
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column   ' xlToLeft
        entetes(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    For Each cle In valeurs.Keys
        If entetes.Exists(CStr(cle)) Then
            ws.Cells(ligne, entetes(CStr(cle))).Value = CStr(valeurs(cle))
        End If
    Next cle
End Sub

Private Function TrouverLigne(ByVal ws As Worksheet, ByVal colonneCle As String, _
                              ByVal valeurCle As String) As Long
    Dim entetes As Object, c As Long, r As Long, colCle As Long, derniere As Long
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column
        entetes(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    If Not entetes.Exists(colonneCle) Then Exit Function
    colCle = entetes(colonneCle)
    derniere = ws.Cells(ws.Rows.Count, colCle).End(-4162).Row
    For r = 2 To derniere
        If Trim$(CStr(ws.Cells(r, colCle).Value)) = valeurCle Then
            TrouverLigne = r
            Exit Function
        End If
    Next r
End Function

Private Function TrouverLignes(ByVal ws As Worksheet, ByVal colonneCle As String, _
                               ByVal valeurCle As String) As Collection
    Dim entetes As Object, c As Long, r As Long, colCle As Long, derniere As Long, res As Collection
    Set res = New Collection
    Set entetes = CreateObject("Scripting.Dictionary")
    entetes.CompareMode = 1
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(-4159).Column
        entetes(Trim$(CStr(ws.Cells(1, c).Value))) = c
    Next c
    Set TrouverLignes = res
    If Not entetes.Exists(colonneCle) Then Exit Function
    colCle = entetes(colonneCle)
    derniere = ws.Cells(ws.Rows.Count, colCle).End(-4162).Row
    For r = 2 To derniere
        If Trim$(CStr(ws.Cells(r, colCle).Value)) = valeurCle Then res.Add r
    Next r
End Function

Private Function ProchainID(ByVal ws As Worksheet, ByVal prefixe As String) As String
    Dim derniere As Long, r As Long, v As String, num As Long, maxi As Long
    derniere = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    For r = 2 To derniere
        v = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(v, Len(prefixe)) = prefixe Then
            num = Val(Mid$(v, Len(prefixe) + 1))
            If num > maxi Then maxi = num
        End If
    Next r
    ProchainID = prefixe & Format$(maxi + 1, "00000")
End Function

' Cree un classeur de base s'il n'existe pas (ex : agenda d'une nouvelle annee)
Public Sub CreerClasseurSiAbsent(ByVal fichier As String, ByVal feuille As String, ByVal entetes As Variant)
    Dim wb As Workbook, i As Long
    If modFichiers.FichierExiste(fichier) Then Exit Sub
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add
    wb.Windows(1).Visible = False
    Do While wb.Worksheets.Count > 1
        Application.DisplayAlerts = False
        wb.Worksheets(wb.Worksheets.Count).Delete
        Application.DisplayAlerts = True
    Loop
    wb.Worksheets(1).Name = feuille
    For i = LBound(entetes) To UBound(entetes)
        wb.Worksheets(1).Cells(1, i - LBound(entetes) + 1).Value = entetes(i)
    Next i
    wb.Worksheets(1).Rows(1).Font.Bold = True
    wb.SaveAs fichier, 51
    wb.Close False
    Application.ScreenUpdating = True
End Sub
