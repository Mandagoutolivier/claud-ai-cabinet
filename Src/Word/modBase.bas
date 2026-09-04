Attribute VB_Name = "modBase"
Option Explicit
' =====================================================================
' modBase - Lecture des bases (poste medecin) SANS verrou :
' copie locale du xlsx puis lecture par une instance Excel INVISIBLE
' (l'ADO/ACE se fige a l'interieur de Word : ne pas y revenir).
' Renvoie des Collections de Scripting.Dictionary (cle = nom de colonne,
' insensible a la casse). Cache memoire de 60 s.
' =====================================================================

Private mPatients As Collection
Private mCorrespondants As Collection
Private mCharge As Date
Private Const CACHE_S As Long = 60

' Lit une feuille d'un classeur (copie locale) et renvoie la Collection.
Public Function LireTable(ByVal fichier As String, ByVal feuille As String, _
                          Optional ByVal colonneNonVide As String = "ID") As Collection
    Dim xl As Object, wb As Object, res As Collection
    Set xl = OuvrirExcel()
    On Error GoTo Nettoyage
    Set wb = xl.Workbooks.Open(modFichiers.CopieLocale(fichier), ReadOnly:=True, AddToMru:=False)
    Set res = LireFeuille(wb, feuille, colonneNonVide)
    wb.Close False
    xl.Quit
    Set LireTable = res
    Exit Function
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    xl.Quit
    On Error GoTo 0
    Err.Raise numErr, "modBase", descErr
End Function

Private Function OuvrirExcel() As Object
    Dim xl As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set OuvrirExcel = xl
End Function

Private Function LireFeuille(ByVal wb As Object, ByVal feuille As String, _
                             ByVal colonneNonVide As String) As Collection
    Dim ws As Object, donnees As Variant, col As Collection, dict As Object
    Dim r As Long, c As Long, nLignes As Long, nCols As Long
    Set ws = wb.Worksheets(feuille)
    Set col = New Collection
    donnees = ws.UsedRange.Value
    If IsEmpty(donnees) Or Not IsArray(donnees) Then
        Set LireFeuille = col
        Exit Function
    End If
    nLignes = UBound(donnees, 1)
    nCols = UBound(donnees, 2)
    For r = 2 To nLignes
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1
        For c = 1 To nCols
            dict(Nz(donnees(1, c))) = Nz(donnees(r, c))
        Next c
        If Len(colonneNonVide) = 0 Then
            col.Add dict
        ElseIf dict.Exists(colonneNonVide) Then
            If Len(dict(colonneNonVide)) > 0 Then col.Add dict
        End If
    Next r
    Set LireFeuille = col
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

' Charge PATIENTS et CORRESPONDANTS en UNE session Excel
Private Sub ChargerBases(ByVal forcer As Boolean)
    Dim xl As Object, wb As Object, tous As Collection, actifs As Collection, c As Object
    If Not forcer And Not mPatients Is Nothing Then
        If DateDiff("s", mCharge, Now) <= CACHE_S Then Exit Sub
    End If
    modLog.Etape "base : ouverture d'Excel invisible"
    Set xl = OuvrirExcel()
    On Error GoTo Nettoyage
    modLog.Etape "base : copie locale de Patients.xlsx"
    Dim copie As String
    copie = modFichiers.CopieLocale(modConfig.FichierPatients())
    modLog.Etape "base : ouverture de la copie " & copie
    Set wb = xl.Workbooks.Open(copie, ReadOnly:=True, AddToMru:=False)
    modLog.Etape "base : lecture des feuilles"
    Set mPatients = LireFeuille(wb, "PATIENTS", "ID")
    Set tous = LireFeuille(wb, "CORRESPONDANTS", "ID")
    wb.Close False
    xl.Quit
    modLog.Etape "base : " & mPatients.Count & " patients charges"
    Set actifs = New Collection
    For Each c In tous
        If c("Actif") <> "0" Then actifs.Add c
    Next c
    Set mCorrespondants = actifs
    mCharge = Now
    Exit Sub
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    xl.Quit
    On Error GoTo 0
    Err.Raise numErr, "modBase", descErr
End Sub

Public Function Patients(Optional ByVal forcer As Boolean = False) As Collection
    ChargerBases forcer
    Set Patients = mPatients
End Function

Public Function Correspondants(Optional ByVal forcer As Boolean = False) As Collection
    ChargerBases forcer
    Set Correspondants = mCorrespondants
End Function

Public Function PatientParID(ByVal id As String) As Object
    Dim p As Object
    For Each p In Patients()
        If p("ID") = id Then Set PatientParID = p: Exit Function
    Next p
    Set PatientParID = Nothing
End Function

Public Function CorrespondantParID(ByVal id As String) As Object
    Dim c As Object
    For Each c In Correspondants()
        If c("ID") = id Then Set CorrespondantParID = c: Exit Function
    Next c
    Set CorrespondantParID = Nothing
End Function
