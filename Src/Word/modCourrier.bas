Attribute VB_Name = "modCourrier"
Option Explicit
' =====================================================================
' modCourrier - Creation des courriers depuis les modeles, remplissage
' de l'en-tete SANS RESSAISIE (patient + correspondant + config), et
' acces au corps du courrier via le signet CORPS.
' Signets attendus dans le modele : EXPEDITEUR, DESTINATAIRE, DATELIEU,
' CONCERNE, APPEL, CORPS, SIGNATURE (voir make_modeles.ps1).
' =====================================================================

' --- Commande principale (Ctrl+Alt+N / commande vocale) --------------
Public Sub NouveauCourrier()
    On Error GoTo Erreur
    Dim pat As Object, cor As Object
    modLog.Etape "NouveauCourrier : lecture de la configuration (" & modConfig.Racine() & ")"
    modLog.Etape "NouveauCourrier : choix du patient"
    Set pat = modPatient.ChoisirPatient()
    If pat Is Nothing Then Exit Sub
    modLog.Etape "NouveauCourrier : choix du destinataire"
    Set cor = modPatient.ChoisirCorrespondant(pat("MedTraitantID"), _
              "Destinataire (Entree = medecin traitant)")
    If cor Is Nothing Then Exit Sub
    modLog.Etape "NouveauCourrier : creation du document"
    CreerCourrierPour pat, cor
    modLog.Etape "NouveauCourrier : termine"
    Exit Sub
Erreur:
    ' capturer l'erreur AVANT toute journalisation (un appel avec gestionnaire
    ' d'erreur reinitialise l'objet Err au retour)
    Dim numErr As Long, descErr As String, srcErr As String, etapeErr As String
    numErr = Err.Number: descErr = Err.Description: srcErr = Err.Source: etapeErr = modLog.DerniereEtape()
    modLog.LogErreur "NouveauCourrier : erreur " & numErr & " (" & srcErr & ") : " & descErr & " | etape : " & etapeErr
    MsgBox "Impossible de creer le courrier :" & vbCrLf & descErr & vbCrLf & vbCrLf & _
           "Etape : " & etapeErr & vbCrLf & "Source : " & srcErr & " (n° " & numErr & ")", vbExclamation, "Cabinet"
End Sub

' --- Creation sans interface (testable, reutilisee par les derivees) --
Public Function CreerCourrierPour(ByVal pat As Object, ByVal cor As Object, _
                                  Optional ByVal typeCourrier As String = "consultation") As Document
    Dim doc As Document
    Set doc = CreerDepuisModele("LETTRE TYPE")
    RemplirEnTete doc, pat, cor
    PreparerStyleCorps doc
    doc.Variables("PatientID") = pat("ID")
    doc.Variables("CorrespondantID") = cor("ID")
    doc.Variables("TypeCourrier") = typeCourrier
    PlacerCurseurCorps doc
    Set CreerCourrierPour = doc
End Function

' nomModele SANS extension : essaie .dotx, .dotm, .dot (le modele reel du
' cabinet peut etre dans n'importe lequel de ces formats)
Private Function CreerDepuisModele(ByVal nomModele As String) As Document
    Dim base As String, ext As Variant, chemin As String
    base = modConfig.Chemin("Modeles") & "\" & nomModele
    For Each ext In Array(".dotx", ".dotm", ".dot")
        chemin = base & ext
        If modFichiers.FichierExiste(chemin) Then
            Set CreerDepuisModele = Documents.Add(Template:=chemin)
            Exit Function
        End If
    Next ext
    Err.Raise vbObjectError + 300, "modCourrier", _
        "Modele introuvable : " & base & " (.dotx/.dotm/.dot)"
End Function

Public Sub RemplirEnTete(ByVal doc As Document, ByVal pat As Object, ByVal cor As Object)
    Dim expediteur As String, destinataire As String, concerne As String
    Dim appel As String, signature As String

    expediteur = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                 modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom") & vbCr & _
                 modConfig.Config("MEDECIN", "Specialite") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne1") & vbCr & _
                 modConfig.Config("MEDECIN", "AdresseLigne2") & vbCr & _
                 "Tel : " & modConfig.Config("MEDECIN", "Telephone")

    ' bloc destinataire pre-compose (classeur des specialistes : structures,
    ' services...) sinon composition classique depuis la fiche
    If cor.Exists("BlocDestinataire") Then destinataire = Trim$(cor("BlocDestinataire"))
    If Len(destinataire) = 0 Then
        destinataire = Trim$(cor("Titre") & " " & cor("Prenom") & " " & cor("Nom"))
        If Len(cor("Specialite")) > 0 Then destinataire = destinataire & vbCr & cor("Specialite")
        If Len(cor("Adresse1")) > 0 Then destinataire = destinataire & vbCr & cor("Adresse1")
        If Len(cor("Adresse2")) > 0 Then destinataire = destinataire & vbCr & cor("Adresse2")
        destinataire = destinataire & vbCr & cor("CP") & " " & cor("Ville")
    End If

    concerne = "Concerne : " & modTexte.CiviliteCourte(pat("Sexe")) & " " & _
               pat("Prenom") & " " & pat("Nom") & ", " & modTexte.NeLe(pat("Sexe")) & " " & pat("DDN")

    appel = cor("FormuleAppel")
    If Len(appel) = 0 Then appel = AppelParDefaut(EstTutoye(cor))

    signature = Replace(modConfig.Config("MEDECIN", "Signature", ""), "|", vbCr)
    If Len(signature) = 0 Then
        signature = modConfig.Config("MEDECIN", "Titre", "Docteur") & " " & _
                    modConfig.Config("MEDECIN", "Prenom") & " " & modConfig.Config("MEDECIN", "Nom")
    End If

    Dim politesse As String
    politesse = cor("FormulePolitesse")
    If Len(politesse) = 0 Then politesse = PolitesseParDefaut(EstTutoye(cor))

    RemplirSignet doc, "EXPEDITEUR", expediteur
    RemplirSignet doc, "DESTINATAIRE", destinataire
    RemplirSignet doc, "DATELIEU", modConfig.Config("GENERAL", "Ville") & ", le " & Format$(Date, "d mmmm yyyy")
    RemplirSignet doc, "CONCERNE", concerne
    RemplirSignet doc, "APPEL", appel
    RemplirSignet doc, "POLITESSE", politesse
    RemplirSignet doc, "SIGNATURE", signature
End Sub

' --- registre du correspondant ----------------------------------------
' Tutoiement : champ Tutoiement / TutoiementVouvoiement de la fiche
' ("tu"), sinon deduit d'une formule d'appel familiere (Cher Ami, Mon Cher...)
Public Function EstTutoye(ByVal cor As Object) As Boolean
    Dim v As String, appel As String
    If cor Is Nothing Then Exit Function
    If cor.Exists("Tutoiement") Then v = LCase$(Trim$(cor("Tutoiement")))
    If Len(v) = 0 And cor.Exists("TutoiementVouvoiement") Then v = LCase$(Trim$(cor("TutoiementVouvoiement")))
    If v = "tu" Then EstTutoye = True: Exit Function
    If v = "vous" Then Exit Function
    If cor.Exists("FormuleAppel") Then appel = modTexte.Plier(cor("FormuleAppel"))
    EstTutoye = (InStr(appel, "ami") > 0 Or InStr(appel, "mon cher") > 0 Or InStr(appel, "ma chere") > 0)
End Function

' Defauts du cabinet (decisions du 04/09/2026) : "Cher Ami" pour les
' correspondants proches (tutoyes), "Cher Confrère" sinon ;
' "Bien cordialement" au tutoiement, "Bien confraternellement" au vouvoiement.
Public Function AppelParDefaut(ByVal tutoiement As Boolean) As String
    If tutoiement Then
        AppelParDefaut = modConfig.Config("COURRIER", "AppelProche", "Cher Ami,")
    Else
        AppelParDefaut = modConfig.Config("COURRIER", "AppelDefaut", "Cher Confrère,")
    End If
End Function

Public Function PolitesseParDefaut(ByVal tutoiement As Boolean) As String
    If tutoiement Then
        PolitesseParDefaut = modConfig.Config("COURRIER", "PolitesseProche", "Bien cordialement.")
    Else
        PolitesseParDefaut = modConfig.Config("COURRIER", "PolitesseDefaut", "Bien confraternellement.")
    End If
End Function

' --- identite du patient dans le corps (Ctrl+Alt+P / voix) -----------
' Le medecin ne dicte JAMAIS l'identite : elle est inseree depuis la base.
Public Sub InsererPatient()
    On Error GoTo Erreur
    Dim pat As Object
    Set pat = modClaude.PatientDuDocument(ActiveDocument)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient.", vbExclamation, "Cabinet"
        Exit Sub
    End If
    Selection.TypeText TexteIdentitePatient(pat)
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Ex : "Monsieur Jean FABREGUE, 91 ans"
Public Function TexteIdentitePatient(ByVal pat As Object) As String
    Dim age As String
    age = CalculerAge(pat("DDN"))
    TexteIdentitePatient = modTexte.Civilite(pat("Sexe")) & " " & pat("Prenom") & " " & pat("Nom") & _
                           IIf(Len(age) > 0, ", " & age & " ans", "")
End Function

Private Function CalculerAge(ByVal ddn As String) As String
    Dim p() As String, naissance As Date, age As Long
    p = Split(Trim$(ddn), "/")
    If UBound(p) <> 2 Then Exit Function
    If Val(p(0)) = 0 Or Val(p(1)) = 0 Or Val(p(2)) = 0 Then Exit Function
    naissance = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
    age = DateDiff("yyyy", naissance, Date)
    If DateSerial(Year(Date), Month(naissance), Day(naissance)) > Date Then age = age - 1
    If age >= 0 And age < 130 Then CalculerAge = CStr(age)
End Function

' Pendant la dictee, chaque Entree cree un paragraphe du style "paragraphe
' suivant" defini dans le modele (chaine corps -> politesse 10 cm ->
' signature 8 cm...). On impose au corps un style dont le suivant est
' lui-meme, avec la mise en forme exacte du premier paragraphe du corps.
Public Sub PreparerStyleCorps(ByVal doc As Document)
    On Error Resume Next
    Dim rng As Range, st As Style, i As Long, n As Long
    Dim formats() As ParagraphFormat
    If Not doc.Bookmarks.Exists("CORPS") Then Exit Sub
    Set rng = doc.Bookmarks("CORPS").Range
    ' 1. photographie de la mise en forme directe de TOUS les paragraphes
    '    (modifier un style peut la faire disparaitre : on la restaurera)
    n = doc.Paragraphs.Count
    ReDim formats(1 To n)
    For i = 1 To n
        Set formats(i) = doc.Paragraphs(i).Format.Duplicate
    Next i
    ' 2. le style du corps (et Normal) s'enchaine sur lui-meme : chaque
    '    Entree pendant la dictee reste un paragraphe de corps
    Set st = doc.Styles(CStr(rng.Paragraphs(1).Style))
    If Not st Is Nothing Then st.NextParagraphStyle = st
    Set st = doc.Styles(wdStyleNormal)
    If Not st Is Nothing Then st.NextParagraphStyle = st
    ' 3. restauration de la mise en forme directe photographiee
    If doc.Paragraphs.Count = n Then
        For i = 1 To n
            doc.Paragraphs(i).Format = formats(i)
        Next i
    End If
    ' 4. espacement des paragraphes du corps : celui des courriers du cabinet
    '    (12 pt avant, 0 apres, interligne 1,15), sans espacement "automatique"
    AppliquerEspacement doc, "APPEL"
    AppliquerEspacement doc, "CORPS"
    AppliquerEspacement doc, "POLITESSE"
    If Err.Number <> 0 Then modLog.LogErreur "PreparerStyleCorps : " & Err.Description
End Sub

' Espacement d'usage des courriers (config [COURRIER] : EspaceAvantPt,
' EspaceApresPt, Interligne) applique aux paragraphes d'un signet.
Private Sub AppliquerEspacement(ByVal doc As Document, ByVal signet As String)
    On Error Resume Next
    Dim p As Paragraph
    If Not doc.Bookmarks.Exists(signet) Then Exit Sub
    For Each p In doc.Bookmarks(signet).Range.Paragraphs
        With p.Format
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .SpaceBefore = modConfig.ConfigNum("COURRIER", "EspaceAvantPt", 12)
            .SpaceAfter = modConfig.ConfigNum("COURRIER", "EspaceApresPt", 0)
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(modConfig.ConfigNum("COURRIER", "Interligne", 1.15))
        End With
    Next p
End Sub

' Remplace le contenu d'un signet et RECREE le signet sur le texte insere
Public Sub RemplirSignet(ByVal doc As Document, ByVal nom As String, ByVal texte As String)
    Dim rng As Range
    If Not doc.Bookmarks.Exists(nom) Then Exit Sub
    Set rng = doc.Bookmarks(nom).Range
    rng.Text = texte
    doc.Bookmarks.Add nom, rng
End Sub

' --- Corps du courrier ------------------------------------------------

Public Function CorpsRange(ByVal doc As Document) As Range
    If doc.Bookmarks.Exists("CORPS") Then
        Set CorpsRange = doc.Bookmarks("CORPS").Range
    ElseIf Selection.Type = wdSelectionNormal And Len(Selection.Range.Text) > 10 Then
        ' repli : la selection manuelle du medecin fait foi
        Set CorpsRange = Selection.Range
    Else
        Err.Raise vbObjectError + 301, "modCourrier", _
            "Signet CORPS introuvable dans ce document et aucune selection : " & _
            "utilisez un courrier cree par 'Nouveau courrier', ou selectionnez le texte a corriger."
    End If
End Function

Public Function RecupererCorps(ByVal doc As Document) As String
    Dim t As String
    t = CorpsRange(doc).Text
    ' retire les marques de paragraphe de tete/queue sans toucher au reste
    Do While Len(t) > 0 And (Right$(t, 1) = vbCr Or Right$(t, 1) = Chr$(7))
        t = Left$(t, Len(t) - 1)
    Loop
    Do While Len(t) > 0 And Left$(t, 1) = vbCr
        t = Mid$(t, 2)
    Loop
    RecupererCorps = t
End Function

Public Sub RemplacerCorps(ByVal doc As Document, ByVal texte As String)
    Dim rng As Range, avaitSignet As Boolean, modele As ParagraphFormat
    avaitSignet = doc.Bookmarks.Exists("CORPS")
    Set rng = CorpsRange(doc)
    ' mise en forme du corps AVANT insertion : Word donnerait sinon au texte
    ' insere celle du paragraphe suivant du modele (espacement automatique...)
    Set modele = rng.Paragraphs(1).Format.Duplicate
    rng.Text = texte & vbCr
    UniformiserParagraphes rng, modele
    If avaitSignet Then doc.Bookmarks.Add "CORPS", rng
    AppliquerEspacement doc, "CORPS"
End Sub

' Impose a tous les paragraphes de la plage la mise en forme de reference
Private Sub UniformiserParagraphes(ByVal rng As Range, ByVal modele As ParagraphFormat)
    On Error Resume Next
    Dim p As Paragraph
    For Each p In rng.Paragraphs
        p.Format = modele
    Next p
End Sub

Public Sub PlacerCurseurCorps(ByVal doc As Document)
    On Error Resume Next
    If doc.Bookmarks.Exists("CORPS") Then
        Dim rng As Range
        Set rng = doc.Bookmarks("CORPS").Range
        rng.Collapse wdCollapseStart
        rng.Move wdCharacter, 1     ' a l'interieur du signet (il s'etend en dictant)
        rng.Select
    End If
End Sub
