Attribute VB_Name = "modClaude"
Option Explicit
' =====================================================================
' modClaude - Appel de l'API Claude (correction et lettres derivees).
'  - cle API : %APPDATA%\CabinetCardio\api.key (poste medecin uniquement)
'  - envoi UTF-8 (ADODB.Stream), reponse analysee par modJson
'  - reessais sur 429/5xx et erreurs reseau
'  - mode Debug (config.ini [API] Debug=1) : le contenu exact envoye est
'    ecrit dans Logs\payload_debug.json (recette sur patients FICTIFS)
' + Orchestration de la commande "CorrigerCourrier" (Ctrl+Alt+C).
' =====================================================================

Private Const URL_API As String = "https://api.anthropic.com/v1/messages"
Private Const URL_OPENAI As String = "https://api.openai.com/v1/responses"

' Fournisseur configure : "claude" (defaut) ou "openai" (config.ini [API] Fournisseur)
Public Function FournisseurApi() As String
    FournisseurApi = LCase$(Trim$(modConfig.Config("API", "Fournisseur", "claude")))
    If FournisseurApi <> "openai" Then FournisseurApi = "claude"
End Function

' Cle OpenAI : variable d'environnement OPENAI_API_KEY (poste du cabinet),
' sinon %APPDATA%\CabinetCardio\openai.key
Public Function LireCleOpenAI() As String
    Dim chemin As String
    LireCleOpenAI = Trim$(Environ$("OPENAI_API_KEY"))
    If Len(LireCleOpenAI) > 0 Then Exit Function
    chemin = Environ$("APPDATA") & "\CabinetCardio\openai.key"
    If modFichiers.FichierExiste(chemin) Then
        LireCleOpenAI = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
    End If
    If Len(LireCleOpenAI) = 0 Then
        Err.Raise vbObjectError + 400, "modClaude", _
            "Cle OpenAI introuvable : definissez la variable d'environnement OPENAI_API_KEY" & vbCrLf & _
            "(ou creez le fichier " & chemin & ")."
    End If
End Function

Public Function LireCleApi() As String
    Dim chemin As String
    chemin = Environ$("APPDATA") & "\CabinetCardio\api.key"
    If Not modFichiers.FichierExiste(chemin) Then
        Err.Raise vbObjectError + 400, "modClaude", _
            "Cle API introuvable : " & chemin & vbCrLf & _
            "Creez ce fichier contenant uniquement votre cle API Anthropic."
    End If
    LireCleApi = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
End Function

' Appel synchrone. Renvoie le texte de la reponse ; leve une erreur claire sinon.
Public Function AppelerClaude(ByVal systeme As String, ByVal utilisateur As String, _
                              Optional ByVal fournisseur As String = "") As String
    Dim json As String, cle As String, essai As Long, statut As Long
    Dim http As Object, reponse As String, racine As Object, texte As String
    Dim attentes As Variant, openai As Boolean, url As String

    If Len(fournisseur) = 0 Then fournisseur = FournisseurApi()
    openai = (LCase$(fournisseur) = "openai")
    If openai Then
        ' OpenAI Responses API (meme API que l'ancien projet R12)
        cle = LireCleOpenAI()
        url = URL_OPENAI
        json = "{""model"":""" & modJson.JsonEchapper(modConfig.Config("API", "ModeleOpenAI", "gpt-4.1")) & """" & _
               ",""max_output_tokens"":" & CLng(modConfig.ConfigNum("API", "MaxTokens", 8000)) & _
               ",""instructions"":""" & modJson.JsonEchapper(systeme) & """" & _
               ",""input"":""" & modJson.JsonEchapper(utilisateur) & """}"
    Else
        cle = LireCleApi()
        url = URL_API
        json = "{""model"":""" & modJson.JsonEchapper(modConfig.Config("API", "Modele", "claude-sonnet-5")) & """" & _
               ",""max_tokens"":" & CLng(modConfig.ConfigNum("API", "MaxTokens", 8000)) & _
               ",""system"":""" & modJson.JsonEchapper(systeme) & """" & _
               ",""messages"":[{""role"":""user"",""content"":""" & modJson.JsonEchapper(utilisateur) & """}]}"
    End If

    If modConfig.ConfigBool("API", "Debug", False) Then
        modFichiers.EnsureDossier modConfig.Chemin("Logs")
        modFichiers.EcrireTexteUTF8 modConfig.Chemin("Logs") & "\payload_debug.json", json
    End If

    attentes = Array(2000, 8000, 20000)
    For essai = 0 To 2
        On Error Resume Next
        Err.Clear
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts 5000, 10000, 30000, CLng(modConfig.ConfigNum("API", "TimeoutReceptionMs", 120000))
        http.Open "POST", url, False
        http.setRequestHeader "content-type", "application/json"
        If openai Then
            http.setRequestHeader "Authorization", "Bearer " & cle
        Else
            http.setRequestHeader "x-api-key", cle
            http.setRequestHeader "anthropic-version", "2023-06-01"
            ' cles "liees a une identite" : le workspace doit etre precise
            Dim workspaceId As String
            workspaceId = modConfig.Config("API", "WorkspaceId", "")
            If Len(workspaceId) > 0 Then http.setRequestHeader "anthropic-workspace-id", workspaceId
        End If
        http.Send TexteVersUtf8(json)
        If Err.Number <> 0 Then
            modLog.LogErreur "API essai " & (essai + 1) & " erreur reseau : " & Err.Description
            On Error GoTo 0
            If essai = 2 Then
                Err.Raise vbObjectError + 401, "modClaude", _
                    "Impossible de joindre l'API (verifiez la connexion internet)."
            End If
            modFichiers.Pause CLng(attentes(essai))
        Else
            On Error GoTo 0
            statut = http.Status
            reponse = http.responseText
            If statut = 200 Then
                Set racine = modJson.JsonParse(reponse)
                If openai Then
                    texte = modJson.JsonTexteReponseOpenAI(racine)
                Else
                    texte = modJson.JsonTexteReponse(racine)
                End If
                If Len(texte) = 0 Then
                    Err.Raise vbObjectError + 402, "modClaude", "Reponse API vide ou inattendue."
                End If
                AppelerClaude = texte
                Exit Function
            ElseIf statut = 401 Then
                Err.Raise vbObjectError + 403, "modClaude", _
                    IIf(openai, "Cle OpenAI refusee (401). Verifiez la variable OPENAI_API_KEY.", _
                                "Cle API refusee (401). Verifiez le fichier api.key.")
            ElseIf statut = 429 Or statut >= 500 Then
                modLog.LogErreur "API essai " & (essai + 1) & " statut " & statut
                If essai = 2 Then
                    Err.Raise vbObjectError + 404, "modClaude", _
                        "API indisponible (statut " & statut & "). Reessayez dans quelques minutes."
                End If
                modFichiers.Pause CLng(attentes(essai))
            Else
                Dim msg As String
                On Error Resume Next
                Set racine = modJson.JsonParse(reponse)
                msg = modJson.JsonErreurMessage(racine)
                On Error GoTo 0
                modLog.LogErreur "API statut " & statut & " : " & msg
                Err.Raise vbObjectError + 405, "modClaude", _
                    "Erreur API (statut " & statut & ") : " & msg
            End If
        End If
    Next essai
End Function

' Chaine -> octets UTF-8 sans BOM (pour http.Send)
Private Function TexteVersUtf8(ByVal s As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                  ' binaire
    st.Position = 3              ' saute le BOM
    TexteVersUtf8 = st.Read(-1)
    st.Close
End Function

' Prompt systeme : fichier Config\prompts\<nom> + courriers de reference
' (avecReferences=False : le marqueur est simplement retire - les lettres
' derivees ne doivent PAS imiter les comptes rendus de consultation)
Public Function ChargerPrompt(ByVal nomFichier As String, _
                              Optional ByVal avecReferences As Boolean = True) As String
    Dim chemin As String, prompt As String, refs As String
    chemin = modConfig.Chemin("Config") & "\prompts\" & nomFichier
    If Not modFichiers.FichierExiste(chemin) Then
        Err.Raise vbObjectError + 406, "modClaude", "Prompt introuvable : " & chemin
    End If
    prompt = modFichiers.LireTexteUTF8(chemin)
    If avecReferences Then refs = ChargerReferences() Else refs = ""
    ChargerPrompt = Replace(prompt, "[COURRIERS_DE_REFERENCE]", refs)
End Function

Private Function ChargerReferences() As String
    Dim dossier As String, f As Variant, refs As String, n As Long, contenu As String
    Const MAX_REFS As Long = 3
    Const MAX_CARS As Long = 5000
    dossier = modConfig.Chemin("Config") & "\style"
    For Each f In modFichiers.ListerFichiers(dossier, ".txt")
        If n >= MAX_REFS Then Exit For
        contenu = modFichiers.LireTexteUTF8(CStr(f))
        If Len(contenu) > MAX_CARS Then contenu = Left$(contenu, MAX_CARS)
        n = n + 1
        refs = refs & vbCrLf & "--- Courrier de référence " & n & " ---" & vbCrLf & contenu & vbCrLf
    Next f
    If n > 0 Then
        ChargerReferences = "Courriers de référence (style à imiter fidèlement) :" & vbCrLf & refs
    Else
        ChargerReferences = ""
    End If
End Function

' =====================================================================
' Commande principale : CORRIGER LE COURRIER (Ctrl+Alt+C / voix)
' =====================================================================
Public Sub CorrigerCourrier()
    On Error GoTo Erreur
    Dim doc As Document, pat As Object, cor As Object, ctx As Object
    Dim corps As String, anonyme As String, problemes As String
    Dim reponse As String, final As String, prog As ufProgression

    Set doc = ActiveDocument
    Set pat = PatientDuDocument(doc)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient (utilisez 'Nouveau courrier').", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    Set cor = CorrespondantDuDocument(doc)

    corps = modCourrier.RecupererCorps(doc)
    If Len(Trim$(corps)) < 10 Then
        MsgBox "Le corps du courrier est vide : dictez d'abord votre texte.", vbExclamation, "Cabinet"
        Exit Sub
    End If

    ' 1. substitutions locales (macros historiques)
    corps = modSubstitutions.AppliquerSubstitutions(corps)

    ' 2. anonymisation dirigee + scan residuel bloquant
    Set ctx = modAnonymise.Construire(pat, cor)
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    problemes = modAnonymise.ScanResiduel(anonyme, ctx)
    If Len(problemes) > 0 Then
        If MsgBox("L'anonymisation a detecte un risque avant envoi :" & vbCrLf & vbCrLf & _
                  problemes & vbCrLf & "Envoyer QUAND MEME a l'API ?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Cabinet - protection des donnees") <> vbYes Then
            Exit Sub
        End If
        modLog.LogInfo "Envoi force malgre scan residuel : " & Replace(problemes, vbCrLf, " / ")
    End If

    ' 3. appel API
    Set prog = New ufProgression
    prog.Show vbModeless
    prog.Definir "Correction du courrier en cours..."
    On Error GoTo ErreurApi
    reponse = AppelerClaude(ChargerPrompt("correction.txt"), anonyme)
    On Error GoTo Erreur
    Unload prog
    Set prog = Nothing

    ' 4. verification des balises au retour
    problemes = modAnonymise.VerifierBalisesRetour(reponse, ctx, anonyme)
    If Len(problemes) > 0 Then
        modFichiers.EcrireTexteUTF8 modConfig.Chemin("Logs") & "\reponse_rejetee.txt", reponse
        MsgBox "La reponse de l'API a altere des balises d'identite ; le texte n'a PAS ete insere." & _
               vbCrLf & problemes, vbCritical, "Cabinet"
        Exit Sub
    End If

    ' 5. reinjection + archivage du brouillon + remplacement (annulable)
    final = modAnonymise.Reinjecter(reponse, ctx)
    ArchiverBrouillon pat, corps
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Correction IA"
    modCourrier.RemplacerCorps doc, NettoyerReponse(final)
    ' 6. mise en gras (medicaments en majuscules, expressions) - dictionnaires Excel
    On Error Resume Next
    Dim nGras As Long
    nGras = modGras.AppliquerGras(doc)
    If Err.Number <> 0 Then modLog.LogErreur "Gras apres correction : " & Err.Description
    On Error GoTo Erreur
    ur.EndCustomRecord
    Application.StatusBar = "Courrier corrige, " & nGras & " mise(s) en gras (Ctrl+Z pour revenir a la dictee brute)."
    Exit Sub

ErreurApi:
    Dim descApi As String
    descApi = Err.Description
    If Not prog Is Nothing Then Unload prog
    MsgBox descApi, vbCritical, "Cabinet - correction impossible"
    Exit Sub
Erreur:
    Dim descErr As String, numErr As Long
    descErr = Err.Description: numErr = Err.Number
    If Not prog Is Nothing Then Unload prog
    modLog.LogErreur "CorrigerCourrier : erreur " & numErr & " : " & descErr
    MsgBox "Erreur : " & descErr, vbCritical, "Cabinet"
End Sub

' Normalisation de la reponse : fins de ligne Word ; les lignes vides entre
' paragraphes (habitude de l'IA) sont supprimees - dans les courriers du
' cabinet, l'espacement de 12 pt separe les paragraphes, jamais une ligne vide.
Public Function NettoyerReponse(ByVal t As String) As String
    Dim lignes() As String, i As Long, res As String, l As String
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    lignes = Split(t, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        l = Trim$(lignes(i))
        If Len(l) > 0 Then
            If Len(res) > 0 Then res = res & vbCr
            res = res & l
        End If
    Next i
    NettoyerReponse = res
End Function

Public Function PatientDuDocument(ByVal doc As Document) As Object
    On Error Resume Next
    Dim id As String
    id = doc.Variables("PatientID")
    On Error GoTo 0
    If Len(id) > 0 Then Set PatientDuDocument = modBase.PatientParID(id)
End Function

Public Function CorrespondantDuDocument(ByVal doc As Document) As Object
    On Error Resume Next
    Dim id As String
    id = doc.Variables("CorrespondantID")
    On Error GoTo 0
    If Len(id) > 0 Then Set CorrespondantDuDocument = modBase.CorrespondantParID(id)
End Function

Private Sub ArchiverBrouillon(ByVal pat As Object, ByVal corps As String)
    On Error Resume Next
    modFichiers.EcrireTexteUTF8 modPatient.DossierPatient(pat) & "\brouillon_" & _
                                modFichiers.IdUnique() & ".txt", corps
End Sub
