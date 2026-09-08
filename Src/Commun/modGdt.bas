Attribute VB_Name = "modGdt"
Option Explicit
' =====================================================================
' modGdt - Envoi de l'identite patient au logiciel ECG Resting12Lead
' par fichier GDT (dialecte valide en essai reel le 31/08/2026) :
'   satz 6302, GDT 02.00, jeu de caracteres 9206=3 (ANSI),
'   date de naissance JJMMAAAA, fichier IMPORT.GDT en ANSI cp1252.
' Cote Resting12Lead : Parametres > Parametres Interface >
'   "Saisie Auto info Patient" = GDT (et "Connection Systeme Info DMS"
'   DEcochee - les deux modes sont exclusifs), Interface GDT In/Out
'   pointant sur le meme dossier que [ECG] DossierGdt.
' =====================================================================

' Ecrit IMPORT.GDT pour ce patient. dossierOverride : pour les tests.
' Renvoie le chemin ecrit. Erreur claire si non configure.
Public Function EcrireGdtPatient(ByVal pat As Object, _
                                 Optional ByVal dossierOverride As String = "") As String
    Dim dossier As String, chemin As String, contenu As String
    dossier = dossierOverride
    If Len(dossier) = 0 Then dossier = modConfig.Config("ECG", "DossierGdt", "")
    If Len(dossier) = 0 Then
        Err.Raise vbObjectError + 800, "modGdt", _
            "Envoi ECG non configure : renseignez [ECG] DossierGdt dans Config\config.ini " & _
            "(dossier surveille par Resting12Lead, ex : C:\Mandagout)."
    End If
    If Right$(dossier, 1) = "\" Then dossier = Left$(dossier, Len(dossier) - 1)
    If Not modFichiers.DossierExiste(dossier) Then
        Err.Raise vbObjectError + 801, "modGdt", "Dossier ECG introuvable : " & dossier
    End If

    contenu = ConstruireGdt(pat)
    chemin = dossier & "\IMPORT.GDT"
    modFichiers.EcrireTexteAnsi chemin, contenu
    modLog.LogInfo "GDT ecrit pour " & pat("ID") & " -> " & chemin
    EcrireGdtPatient = chemin
End Function

Public Function ConstruireGdt(ByVal pat As Object) As String
    Dim lignes As Collection, l As Variant, total As Long, contenu As String
    Set lignes = New Collection
    lignes.Add LigneGdt("8000", "6302")
    lignes.Add "PLACEHOLDER"
    lignes.Add LigneGdt("9206", "3")
    lignes.Add LigneGdt("9218", "02.00")
    lignes.Add LigneGdt("3000", pat("ID"))
    lignes.Add LigneGdt("3101", UCase$(pat("Nom")))
    lignes.Add LigneGdt("3102", pat("Prenom"))
    Dim ddn As String, sexe As String
    ddn = DdnVersGdt(modTexte.DdnPatient(pat))
    If Len(ddn) > 0 Then
        lignes.Add LigneGdt("3103", ddn)
    Else
        modLog.LogErreur "GDT " & pat("ID") & " : date de naissance absente ou illisible ('" & modTexte.DdnPatient(pat) & "') - a completer dans la fiche"
    End If
    ' 3110 sexe (1 = masculin, 2 = feminin) : utile aux normes ECG, et donne
    ' au logiciel ECG une fiche complete des l'import
    sexe = SexeVersGdt(pat)
    If Len(sexe) > 0 Then
        lignes.Add LigneGdt("3110", sexe)
    Else
        modLog.LogErreur "GDT " & pat("ID") & " : sexe absent ou illisible - a completer dans la fiche"
    End If
    lignes.Add LigneGdt("8402", modConfig.Config("ECG", "CodeExamen", "EKG01"))

    ' champ 8100 = longueur totale, sa propre ligne comprise (14 octets)
    total = 14
    For Each l In lignes
        If l <> "PLACEHOLDER" Then total = total + Len(l) + 2
    Next l

    For Each l In lignes
        If l = "PLACEHOLDER" Then
            contenu = contenu & LigneGdt("8100", Format$(total, "00000")) & vbCrLf
        Else
            contenu = contenu & l & vbCrLf
        End If
    Next l
    ConstruireGdt = contenu
End Function

Private Function LigneGdt(ByVal champ As String, ByVal valeur As String) As String
    LigneGdt = Format$(Len(champ & valeur) + 5, "000") & champ & valeur
End Function

' Date de naissance -> JJMMAAAA. Tolere "01/01/1935", "1/1/35", "1935-01-01",
' "01.01.1935" et un numero de serie Excel ; "" si illisible.
Private Function DdnVersGdt(ByVal ddn As String) As String
    Dim p() As String, t As String, d As Date, a As Long
    t = Trim$(ddn)
    If Len(t) = 0 Then Exit Function
    t = Replace(Replace(t, ".", "/"), "-", "/")
    p = Split(t, "/")
    On Error Resume Next
    If UBound(p) = 2 Then
        If Len(p(0)) = 4 Then                     ' AAAA/MM/JJ
            d = DateSerial(Val(p(0)), Val(p(1)), Val(p(2)))
        Else                                      ' JJ/MM/AAAA ou JJ/MM/AA
            a = Val(p(2))
            If a < 100 Then a = a + IIf(a > Year(Date) Mod 100, 1900, 2000)
            d = DateSerial(a, Val(p(1)), Val(p(0)))
        End If
    ElseIf IsNumeric(t) Then                      ' numero de serie Excel
        d = CDate(Val(t))
    ElseIf IsDate(t) Then
        d = CDate(t)
    End If
    On Error GoTo 0
    If d = 0 Or Year(d) < 1880 Or d > Date Then Exit Function
    DdnVersGdt = Format$(d, "ddmmyyyy")
End Function

' Sexe de la fiche -> code GDT : M/H/1/Masculin/Homme -> 1, F/2/Feminin/Femme -> 2
Private Function SexeVersGdt(ByVal pat As Object) As String
    Select Case modTexte.SexePatient(pat)
        Case "M": SexeVersGdt = "1"
        Case "F": SexeVersGdt = "2"
    End Select
End Function
