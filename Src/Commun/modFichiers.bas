Attribute VB_Name = "modFichiers"
Option Explicit
' =====================================================================
' modFichiers - Acces fichiers robuste :
'  - lecture/ecriture texte UTF-8 (ADODB.Stream)
'  - copie locale des bases avec reessais (lecture sans verrou)
'  - verrous cooperatifs (Base\locks\*.lock) pour les ecritures courtes
'  - fichiers-drapeaux (Echange\) : ecriture atomique .tmp -> .txt
'  - sauvegardes horodatees avec purge
' =====================================================================

Private Const VERROU_PERIME_S As Long = 120   ' verrou abandonne apres crash

Public Sub Pause(ByVal ms As Long)
    Dim fin As Single
    fin = Timer + ms / 1000!
    If fin >= 86400! Then fin = fin - 86400!   ' passage minuit : au pire pause courte
    Do While Timer < fin And fin - Timer < 60!
        DoEvents
    Loop
End Sub

' Tests d'existence par FileSystemObject : Dir$() s'est revele non fiable
' dans certaines sessions Word interactives (retour vide sur un fichier present).
Public Function FichierExiste(ByVal chemin As String) As Boolean
    On Error Resume Next
    FichierExiste = CreateObject("Scripting.FileSystemObject").FileExists(chemin)
End Function

Public Function DossierExiste(ByVal chemin As String) As Boolean
    On Error Resume Next
    DossierExiste = CreateObject("Scripting.FileSystemObject").FolderExists(chemin)
End Function

' Liste des fichiers d'un dossier correspondant a une extension (".txt"),
' tries par nom ; Collection de chemins complets.
Public Function ListerFichiers(ByVal dossier As String, ByVal extension As String) As Collection
    Dim fso As Object, f As Object, col As Collection, n As Variant
    Set col = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dossier) Then Set ListerFichiers = col: Exit Function
    For Each f In fso.GetFolder(dossier).Files
        If LCase$(Right$(f.Name, Len(extension))) = LCase$(extension) Then InsererTrie col, f.Name
    Next f
    Dim res As Collection
    Set res = New Collection
    For Each n In col
        res.Add dossier & "\" & n
    Next n
    Set ListerFichiers = res
End Function

' Insertion d'une chaine dans une Collection maintenue triee (VBA pur,
' sans composant .NET : System.Collections.ArrayList echoue dans Office)
Private Sub InsererTrie(ByVal col As Collection, ByVal valeur As String)
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col(i), valeur, vbTextCompare) > 0 Then
            col.Add valeur, , i
            Exit Sub
        End If
    Next i
    col.Add valeur
End Sub

Public Sub EnsureDossier(ByVal chemin As String)
    Dim fso As Object, parties() As String, i As Long, courant As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(chemin) Then Exit Sub
    parties = Split(chemin, "\")
    courant = parties(0)                       ' "C:" ou "" (UNC)
    If Len(courant) = 0 Then courant = "\\" & parties(2) & "\" & parties(3): i = 4 Else i = 1
    Do While i <= UBound(parties)
        courant = courant & "\" & parties(i)
        If Not fso.FolderExists(courant) Then fso.CreateFolder courant
        i = i + 1
    Loop
End Sub

' --- Texte UTF-8 -----------------------------------------------------

Public Function LireTexteUTF8(ByVal chemin As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                ' texte
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile chemin
    LireTexteUTF8 = st.ReadText(-1)
    st.Close
    ' retirer un eventuel BOM residuel
    If Len(LireTexteUTF8) > 0 Then
        If AscW(Left$(LireTexteUTF8, 1)) = &HFEFF Then LireTexteUTF8 = Mid$(LireTexteUTF8, 2)
    End If
End Function

' Ecriture ANSI (cp1252) pour les logiciels anciens (fichiers GDT...)
Public Sub EcrireTexteAnsi(ByVal chemin As String, ByVal contenu As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(chemin, True, False)   ' False = ANSI
    ts.Write contenu
    ts.Close
End Sub

Public Sub EcrireTexteUTF8(ByVal chemin As String, ByVal contenu As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText contenu
    st.SaveToFile chemin, 2    ' ecrase
    st.Close
End Sub

' --- Copie locale (lecture des bases sans verrou) --------------------

Public Function DossierTempLocal() As String
    DossierTempLocal = Environ$("TEMP") & "\CabinetCardio"
    EnsureDossier DossierTempLocal
End Function

' Copie <chemin> vers %TEMP%\CabinetCardio\ et renvoie le chemin local.
' Reessaie si le fichier est brievement verrouille par une ecriture.
Public Function CopieLocale(ByVal chemin As String, Optional ByVal essais As Long = 4, _
                            Optional ByVal delaiMs As Long = 500) As String
    Dim fso As Object, dest As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(chemin) Then
        Err.Raise vbObjectError + 110, "modFichiers", "Fichier introuvable : " & chemin
    End If
    dest = DossierTempLocal() & "\" & fso.GetFileName(chemin)
    For i = 1 To essais
        On Error Resume Next
        Err.Clear
        fso.CopyFile chemin, dest, True
        If Err.Number = 0 Then
            On Error GoTo 0
            CopieLocale = dest
            Exit Function
        End If
        On Error GoTo 0
        Pause delaiMs
    Next i
    Err.Raise vbObjectError + 111, "modFichiers", _
        "Impossible de copier " & chemin & " (fichier occupe apres " & essais & " tentatives)."
End Function

' --- Verrous cooperatifs ---------------------------------------------

Private Function CheminVerrou(ByVal nom As String) As String
    Dim d As String
    d = modConfig.Chemin("Base") & "\locks"
    EnsureDossier d
    CheminVerrou = d & "\" & nom & ".lock"
End Function

Public Function AcquerirVerrou(ByVal nom As String, Optional ByVal timeoutMs As Long = 5000) As Boolean
    Dim fso As Object, chemin As String, debut As Single, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    chemin = CheminVerrou(nom)
    debut = Timer
    Do
        ' verrou perime (poste plante) : on le purge
        If fso.FileExists(chemin) Then
            On Error Resume Next
            If DateDiff("s", fso.GetFile(chemin).DateLastModified, Now) > VERROU_PERIME_S Then
                fso.DeleteFile chemin, True
                modLog.LogInfo "Verrou perime purge : " & nom
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        Err.Clear
        Set ts = fso.CreateTextFile(chemin, False)   ' echoue si deja pris
        If Err.Number = 0 Then
            ts.WriteLine Environ$("COMPUTERNAME") & " " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
            ts.Close
            On Error GoTo 0
            AcquerirVerrou = True
            Exit Function
        End If
        On Error GoTo 0
        Pause 400
    Loop While Timer - debut < timeoutMs / 1000!
    AcquerirVerrou = False
End Function

Public Sub RelacherVerrou(ByVal nom As String)
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject").DeleteFile CheminVerrou(nom), True
End Sub

' --- Fichiers-drapeaux (communication medecin -> secretaire) ---------

' dict : Scripting.Dictionary cle->valeur (valeurs sans saut de ligne).
' Ecrit dans <dossier>\<nomBase>.txt de facon atomique (.tmp puis renommage).
Public Function EcrireDrapeau(ByVal dossier As String, ByVal nomBase As String, ByVal dict As Object) As String
    Dim contenu As String, cle As Variant, fso As Object
    Dim tmpChemin As String, finalChemin As String
    EnsureDossier dossier
    For Each cle In dict.Keys
        contenu = contenu & cle & "=" & Replace(Replace(CStr(dict(cle)), vbCr, " "), vbLf, " ") & vbCrLf
    Next cle
    tmpChemin = dossier & "\" & nomBase & ".tmp"
    finalChemin = dossier & "\" & nomBase & ".txt"
    EcrireTexteUTF8 tmpChemin, contenu
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(finalChemin) Then fso.DeleteFile finalChemin, True
    fso.MoveFile tmpChemin, finalChemin
    EcrireDrapeau = finalChemin
End Function

Public Function LireDrapeau(ByVal chemin As String) As Object
    Dim dict As Object, lignes() As String, i As Long, p As Long, contenu As String
    Set dict = CreateObject("Scripting.Dictionary")
    contenu = LireTexteUTF8(chemin)
    contenu = Replace(contenu, vbCrLf, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        p = InStr(lignes(i), "=")
        If p > 0 Then dict(Left$(lignes(i), p - 1)) = Mid$(lignes(i), p + 1)
    Next i
    Set LireDrapeau = dict
End Function

' Identifiant unique lisible pour nommer drapeaux et fichiers
Public Function IdUnique() As String
    Randomize
    IdUnique = Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int(Rnd * 100000), "00000")
End Function

' --- Sauvegardes horodatees ------------------------------------------

Public Sub SauvegardeHorodatee(ByVal chemin As String, Optional ByVal maxVersions As Long = 60)
    On Error Resume Next
    Dim fso As Object, dossierSvg As String, nomBase As String, ext As String, dest As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(chemin) Then Exit Sub
    dossierSvg = modConfig.Chemin("Sauvegardes")
    EnsureDossier dossierSvg
    nomBase = fso.GetBaseName(chemin)
    ext = fso.GetExtensionName(chemin)
    dest = dossierSvg & "\" & nomBase & "_" & Format$(Now, "yyyymmdd_hhnnss") & "." & ext
    fso.CopyFile chemin, dest, True
    PurgerSauvegardes dossierSvg, nomBase & "_", maxVersions
End Sub

Private Sub PurgerSauvegardes(ByVal dossier As String, ByVal prefixe As String, ByVal maxVersions As Long)
    On Error Resume Next
    Dim fso As Object, f As Object, noms As Collection, n As Long, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set noms = New Collection
    For Each f In fso.GetFolder(dossier).Files
        If LCase$(Left$(f.Name, Len(prefixe))) = LCase$(prefixe) Then InsererTrie noms, f.Name
    Next f
    n = noms.Count - maxVersions               ' horodatage lexical = chronologique
    For i = 1 To n
        fso.DeleteFile dossier & "\" & noms(i), True
    Next i
End Sub

' Nettoie une chaine pour un nom de fichier/dossier Windows
Public Function NomFichierSur(ByVal s As String) As String
    Dim interdits As Variant, c As Variant
    interdits = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab)
    For Each c In interdits
        s = Replace(s, CStr(c), " ")
    Next c
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NomFichierSur = Trim$(s)
End Function
