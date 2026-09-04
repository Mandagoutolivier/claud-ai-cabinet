Attribute VB_Name = "modRaccourcis"
Option Explicit
' =====================================================================
' modRaccourcis - Aide et reinstallation des raccourcis clavier.
' Les raccourcis sont normalement enregistres dans Cabinet.dotm par le
' build ; cette macro les repose au besoin, et AideCabinet les rappelle.
' Les commandes vocales Dragon envoient ces memes raccourcis.
' =====================================================================

Public Sub AideCabinet()
    MsgBox "Commandes du cabinet :" & vbCrLf & vbCrLf & _
           "Ctrl+Alt+N  -  Nouveau courrier (choix du patient)" & vbCrLf & _
           "Ctrl+Alt+P  -  Inserer l'identite du patient (nom + age) au curseur" & vbCrLf & _
           "Ctrl+Alt+C  -  Corriger le courrier dicte (IA)" & vbCrLf & _
           "Ctrl+Alt+D  -  Lettre derivee (demande d'examen ou d'avis)" & vbCrLf & _
           "Ctrl+Alt+G  -  Envoyer l'identite au poste ECG (Resting12Lead)" & vbCrLf & _
           "Ctrl+Alt+V  -  Valider et transmettre au secretariat" & vbCrLf & vbCrLf & _
           "Commandes vocales Dragon : 'nouveau courrier', 'insere le patient'," & vbCrLf & _
           "'corrige le courrier', 'lettre derivee', 'valide le courrier'.", _
           vbInformation, "Cabinet - aide"
End Sub

' Execute automatiquement au chargement de Cabinet.dotm (demarrage de Word) :
' verifie et repare les raccourcis sans jamais demander d'enregistrer le modele.
Public Sub AutoExec()
    On Error Resume Next
    modLog.LogInfo "AutoExec Cabinet.dotm : demarrage"
    InstallerRaccourcisSession False
End Sub

' Sonde sans interface : prouve qu'une liaison de touche declenche bien une
' macro (ecrit un fichier temoin). Utilisee par les tests, jamais par l'utilisateur.
Public Sub SondeRaccourci()
    On Error Resume Next
    modFichiers.EnsureDossier Environ$("TEMP") & "\CabinetCardio"
    modFichiers.EcrireTexteAnsi Environ$("TEMP") & "\CabinetCardio\sonde_raccourci.txt", Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub ReinstallerRaccourcis()
    InstallerRaccourcisSession True
End Sub

Private Sub InstallerRaccourcisSession(ByVal verbeux As Boolean)
    On Error Resume Next
    Dim modele As Template, i As Long, repares As Long
    Dim touches As Variant, macros As Variant
    Set modele = TrouverModeleCabinet()
    If modele Is Nothing Then
        If verbeux Then MsgBox "Cabinet.dotm n'est pas charge.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    ' noms NON qualifies : c'est la forme que Word resout pour un modele global
    touches = Array(wdKeyN, wdKeyC, wdKeyD, wdKeyP, wdKeyG, wdKeyV, wdKeyB)
    macros = Array("NouveauCourrier", "CorrigerCourrier", "LettreDerivee", _
                   "InsererPatient", "EnvoyerECG", "ValiderCourrier", "MettreEnGras")
    CustomizationContext = modele
    For i = 0 To UBound(touches)
        Dim code As Long
        code = BuildKeyCode(wdKeyControl, wdKeyAlt, touches(i))
        ' on repose systematiquement la liaison (la propriete Command d'un
        ' modele global se lit vide, elle ne permet pas de verifier)
        Err.Clear
        KeyBindings.Add wdKeyCategoryMacro, macros(i), code
        If Err.Number = 0 Then repares = repares + 1 Else modLog.LogErreur "raccourci " & macros(i) & " : " & Err.Description
    Next i
    modele.Saved = True          ' aucune invite d'enregistrement du modele
    If verbeux Then MsgBox "Raccourcis verifies : " & repares & " reinstalle(s) (Ctrl+Alt+N/C/D/P/G/V/B).", vbInformation, "Cabinet"
    modLog.LogInfo "Raccourcis (re)poses : " & repares & "/7"
End Sub

' Diagnostic a lancer depuis Alt+F8 en cas de probleme : etat de la configuration du poste
Public Sub DiagnosticCabinet()
    On Error Resume Next
    Dim r As String, f As String, racine As String, n As Long
    f = Environ$("APPDATA") & "\CabinetCardio\chemin.txt"
    r = "APPDATA : " & Environ$("APPDATA") & vbCrLf
    r = r & "chemin.txt (FSO) : " & modFichiers.FichierExiste(f) & "   (Dir$) : " & (Len(Dir$(f)) > 0) & vbCrLf
    Err.Clear
    racine = modConfig.Racine()
    If Err.Number <> 0 Then r = r & "Racine : ERREUR " & Err.Description & vbCrLf: Err.Clear Else r = r & "Racine : " & racine & vbCrLf
    r = r & "Racine existe (FSO) : " & modFichiers.DossierExiste(racine) & vbCrLf
    r = r & "config.ini modele API : " & modConfig.Config("API", "Modele", "(absent)") & vbCrLf
    r = r & "Patients.xlsx : " & modFichiers.FichierExiste(modConfig.FichierPatients()) & vbCrLf
    r = r & "LETTRE TYPE.dot : " & modFichiers.FichierExiste(modConfig.Chemin("Modeles") & "\LETTRE TYPE.dot") & vbCrLf
    r = r & "api.key : " & modFichiers.FichierExiste(Environ$("APPDATA") & "\CabinetCardio\api.key") & vbCrLf
    Err.Clear
    n = modBase.Patients(True).Count
    If Err.Number <> 0 Then r = r & "Lecture base : ERREUR " & Err.Description & vbCrLf: Err.Clear Else r = r & "Lecture base : " & n & " patients" & vbCrLf
    Dim t As Template
    For Each t In Templates
        r = r & "Modele charge : " & t.Name & vbCrLf
    Next t
    modLog.LogInfo "Diagnostic : " & Replace(r, vbCrLf, " | ")
    MsgBox r, vbInformation, "Cabinet - diagnostic"
End Sub

Private Function TrouverModeleCabinet() As Template
    Dim t As Template
    For Each t In Templates
        If LCase$(t.Name) = "cabinet.dotm" Then
            Set TrouverModeleCabinet = t
            Exit Function
        End If
    Next t
End Function
