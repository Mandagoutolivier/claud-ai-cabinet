Attribute VB_Name = "modDerivees"
Option Explicit
' =====================================================================
' modDerivees - Lettres derivees du courrier dicte : DEMANDE d'examen
' (épreuve d'effort, scintigraphie, scanner coronaire, IRM cardiaque),
' d'avis specialise ou d'hospitalisation, adressee a un confrere.
'
' Une lettre derivee n'est PAS un resume du compte rendu : c'est une
' demande courte (voir LETTRES_DERIVEES.md). Le corps est redige par
' l'API a partir du corps ANONYMISE du courrier source, avec des
' consignes propres a chaque type de demande et a chaque modalité.
' L'en-tete, l'appel et la formule finale viennent du modele et de la
' fiche correspondant (modCourrier) : l'API ne redige QUE le corps.
'
' Le destinataire est pre-rempli depuis le classeur des correspondants
' specialistes ([DERIVEES] FichierSpecialistes, feuilles Specialistes +
' Specialistes_ParType), filtre sur le type d'examen et trie par
' priorite ; a defaut, la liste generale des correspondants.
' =====================================================================

' Commande principale (Ctrl+Alt+D / commande vocale)
Public Sub LettreDerivee()
    On Error GoTo Erreur
    Dim doc As Document, pat As Object, dest As Object
    Dim code As String, modalite As String, motif As String

    Set doc = ActiveDocument
    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient (utilisez 'Nouveau courrier').", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If

    code = ChoisirTypeDemande()
    If Len(code) = 0 Then Exit Sub

    modalite = ChoisirModalite(code)
    If modalite = vbNullChar Then Exit Sub          ' annulation

    motif = Trim$(InputBox("Motif ou question posée (facultatif, une phrase) :", _
                           "Demande : " & LibelleDemande(code, modalite)))

    Set dest = ChoisirDestinataire(code, modalite)
    If dest Is Nothing Then Exit Sub

    GenererDerivee doc, code, dest, modalite, motif
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    modLog.LogErreur "LettreDerivee : " & descErr
    MsgBox "Erreur : " & descErr, vbCritical, "Cabinet"
End Sub

' ---------------------------------------------------------------------
' Catalogue des demandes (codes stables, utilises par config.ini
' [DERIVEES] Types=TE;SCINTI;SCANCORO;IRM;AVIS;HOSP)
' ---------------------------------------------------------------------
Private Function Catalogue() As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Definir("TE", "Épreuve d'effort", "TEST_EFFORT", "", _
        "Examen demandé : une épreuve d'effort. Première phrase : ""Merci de réaliser une épreuve d'effort à <patient>."" " & _
        "Dernière phrase : ""Je <te/vous> serais reconnaissant de bien vouloir réaliser cet examen afin de compléter ce bilan."""), "TE"
    c.Add Definir("SCINTI", "Scintigraphie myocardique", "SCINTIGRAPHIE_MYOCARDIQUE", _
        "d'effort=une scintigraphie myocardique de perfusion d'effort|" & _
        "sous stress pharmacologique=une scintigraphie myocardique de perfusion sous stress pharmacologique (dipyridamole ou régadénoson)|" & _
        "couplée effort + pharmacologique=une scintigraphie myocardique de perfusion sous effort couplé à un stress pharmacologique|" & _
        "de repos=une scintigraphie myocardique de perfusion de repos", _
        "Examen demandé : <modalite>. Première phrase : ""Merci de réaliser <modalite> à <patient>."" " & _
        "Si le courrier source mentionne une contre-indication à l'effort, un bloc de branche gauche ou un stimulateur, le rappeler en une ligne car cela justifie la modalité. " & _
        "Dernière phrase : ""Je <te/vous> serais reconnaissant de bien vouloir réaliser cet examen et de m'adresser le compte rendu."""), "SCINTI"
    c.Add Definir("SCANCORO", "Scanner coronaire", "COROSCANNER,SCORE_CALCIQUE", _
        "avec score calcique=un scanner coronaire avec score calcique|" & _
        "score calcique seul=un score calcique coronaire|" & _
        "sans score calcique=un scanner coronaire", _
        "Examen demandé : <modalite>. Première phrase : ""Merci de réaliser <modalite> à <patient>."" " & _
        "Si la fonction rénale (créatinine, clairance) figure dans le courrier source, la donner en une ligne ; sinon ne rien inventer. " & _
        "Dernière phrase : ""Je <te/vous> serais reconnaissant de bien vouloir réaliser cet examen."""), "SCANCORO"
    c.Add Definir("IRM", "IRM cardiaque", "IRM", _
        "de repos=une IRM cardiaque de repos|" & _
        "de stress sous dobutamine=une IRM cardiaque de stress sous dobutamine|" & _
        "de stress sous vasodilatateur=une IRM cardiaque de stress sous vasodilatateur (adénosine ou régadénoson)", _
        "Examen demandé : <modalite>. Première phrase : ""Merci de réaliser <modalite> à <patient>."" " & _
        "Reprendre en une ligne le résultat échographique utile s'il figure dans le courrier source (FEVG, cinétique, valvulopathie). " & _
        "Formuler la question posée en une phrase (viabilité, ischémie, cardiomyopathie...) d'après le motif ou le courrier source. " & _
        "Dernière phrase : ""Je <te/vous> serais reconnaissant de bien vouloir réaliser cet examen."""), "IRM"
    c.Add Definir("AVIS", "Avis spécialisé", "AVIS_NEUROLOGIE,AVIS_GASTRO_ENTEROLOGIE,SAOS,AVIS_PNEUMOLOGIE,AVIS_RYTHMOLOGIE,AVIS_DIABETOLOGIE,AVIS_VASCULAIRE,AVIS_AUTRE", _
        "neurologique=un avis neurologique;AVIS_NEUROLOGIE|" & _
        "gastro-entérologique=un avis gastro-entérologique;AVIS_GASTRO_ENTEROLOGIE|" & _
        "SAOS (apnées du sommeil)=un avis pneumologique à la recherche d'un syndrome d'apnées du sommeil;SAOS,AVIS_PNEUMOLOGIE|" & _
        "rythmologique=un avis rythmologique;AVIS_RYTHMOLOGIE|" & _
        "diabétologique=un avis diabétologique;AVIS_DIABETOLOGIE|" & _
        "vasculaire=un avis vasculaire;AVIS_VASCULAIRE|" & _
        "autre spécialité=un avis spécialisé;AVIS_AUTRE", _
        "Avis demandé : <modalite>. Première phrase : ""Je <te/vous> confie <patient> pour <modalite>."" " & _
        "Puis le motif de l'avis (une ou deux lignes), les symptômes et le traitement en cours s'ils figurent dans le courrier source, et la question posée en une phrase. " & _
        "Dernière phrase : ""Je <te/vous> serais reconnaissant de bien vouloir <recevoir ce patient / cette patiente>."""), "AVIS"
    c.Add Definir("HOSP", "Demande d'hospitalisation", "HOSPITALISATION", _
        "programmée=une hospitalisation programmée|" & _
        "rapide (sous 48 h)=une hospitalisation rapide, sous 48 heures|" & _
        "en urgence=une hospitalisation en urgence", _
        "Demande : <modalite>. Première phrase : ""Je <te/vous> adresse <patient> pour <modalite> dans <ton/votre> service."" " & _
        "Puis le motif d'hospitalisation, les symptômes, l'ECG, l'échographie, le traitement en cours et les antécédents cardiologiques majeurs, uniquement s'ils figurent dans le courrier source, chacun en une ligne. " & _
        "Terminer par la question posée ou l'objectif de l'hospitalisation, puis : ""Je <te/vous> remercie de bien vouloir <prendre en charge ce patient / cette patiente>."""), "HOSP"
    Set Catalogue = c
End Function

Private Function Definir(ByVal code As String, ByVal libelle As String, ByVal typesExamen As String, _
                     ByVal modalites As String, ByVal consigne As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = code
    d("Code") = code
    d("Libelle") = libelle
    d("TypesExamen") = typesExamen
    d("Modalites") = modalites
    d("ConsigneType") = consigne
    Set Definir = d
End Function

Public Function DefinitionDemande(ByVal code As String) As Object
    Dim c As Collection
    Set c = Catalogue()
    On Error Resume Next
    Set DefinitionDemande = c(code)
    On Error GoTo 0
End Function

' Types actives : config.ini [DERIVEES] Types=TE;SCINTI;... (codes du catalogue)
Private Function ChoisirTypeDemande() As String
    Dim codes() As String, i As Long, items As Collection, d As Object, f As ufListe
    codes = Split(modConfig.Config("DERIVEES", "Types", "TE;SCINTI;SCANCORO;IRM;AVIS;HOSP"), ";")
    Set items = New Collection
    For i = LBound(codes) To UBound(codes)
        Set d = DefinitionDemande(UCase$(Trim$(codes(i))))
        If Not d Is Nothing Then items.Add d
    Next i
    If items.Count = 0 Then
        Err.Raise vbObjectError + 502, "modDerivees", _
            "Aucun type de demande valide dans config.ini [DERIVEES] Types (codes : TE;SCINTI;SCANCORO;IRM;AVIS;HOSP)."
    End If
    Set f = New ufListe
    f.Configurer "Type de demande", items, Array("Libelle"), "280 pt"
    f.Show vbModal
    If Not f.Annule Then ChoisirTypeDemande = f.Resultat("Code")
    Unload f
End Function

' Modalite de la demande (libelle affiche). "" si le type n'en a pas,
' vbNullChar si le medecin annule.
Private Function ChoisirModalite(ByVal code As String) As String
    Dim d As Object, mods() As String, i As Long, items As Collection, it As Object, f As ufListe
    Set d = DefinitionDemande(code)
    If Len(d("Modalites")) = 0 Then ChoisirModalite = "": Exit Function
    mods = Split(d("Modalites"), "|")
    Set items = New Collection
    For i = LBound(mods) To UBound(mods)
        Set it = CreateObject("Scripting.Dictionary")
        it.CompareMode = 1
        it("ID") = CStr(i)
        it("Libelle") = Left$(mods(i), InStr(mods(i), "=") - 1)
        items.Add it
    Next i
    Set f = New ufListe
    f.Configurer d("Libelle") & " : modalite", items, Array("Libelle"), "280 pt"
    f.Show vbModal
    If f.Annule Then ChoisirModalite = vbNullChar Else ChoisirModalite = f.Resultat("Libelle")
    Unload f
End Function

' Texte de la modalité pour l'API ("une scintigraphie ... d'effort") et,
' pour les avis, les codes TypeExamen apres le ";".
Private Function DetailModalite(ByVal code As String, ByVal modalite As String) As String
    Dim d As Object, mods() As String, i As Long, p As Long
    Set d = DefinitionDemande(code)
    If Len(modalite) = 0 Or Len(d("Modalites")) = 0 Then Exit Function
    mods = Split(d("Modalites"), "|")
    For i = LBound(mods) To UBound(mods)
        p = InStr(mods(i), "=")
        If StrComp(Left$(mods(i), p - 1), modalite, vbTextCompare) = 0 Then
            DetailModalite = Mid$(mods(i), p + 1)
            Exit Function
        End If
    Next i
End Function

Private Function TexteModalite(ByVal code As String, ByVal modalite As String) As String
    Dim det As String
    det = DetailModalite(code, modalite)
    If InStr(det, ";") > 0 Then det = Left$(det, InStr(det, ";") - 1)
    TexteModalite = det
End Function

' Codes TypeExamen a proposer pour le destinataire (les avis dependent de la specialite)
Public Function TypesExamenPour(ByVal code As String, ByVal modalite As String) As String
    Dim d As Object, det As String
    Set d = DefinitionDemande(code)
    det = DetailModalite(code, modalite)
    If InStr(det, ";") > 0 Then
        TypesExamenPour = Mid$(det, InStr(det, ";") + 1)
    Else
        TypesExamenPour = d("TypesExamen")
    End If
End Function

Public Function LibelleDemande(ByVal code As String, ByVal modalite As String) As String
    Dim d As Object
    Set d = DefinitionDemande(code)
    LibelleDemande = d("Libelle")
    If Len(modalite) > 0 Then LibelleDemande = LibelleDemande & " " & modalite
End Function

' ---------------------------------------------------------------------
' Destinataire pre-rempli depuis le classeur des specialistes
' ---------------------------------------------------------------------
Private Function ChoisirDestinataire(ByVal code As String, ByVal modalite As String) As Object
    Dim items As Collection, f As ufListe, titre As String
    titre = "Destinataire : " & LibelleDemande(code, modalite)
    Set items = CorrespondantsPourTypes(TypesExamenPour(code, modalite))
    If items.Count > 0 Then
        Set f = New ufListe
        f.Configurer titre & "  (Nouveau... = autre correspondant)", items, _
                     Array("NomDestinataire", "Structure", "Ville"), "160 pt;150 pt;80 pt", "", True
        f.Show vbModal
        If Not f.Annule Then
            Set ChoisirDestinataire = f.Resultat
            Unload f
            Exit Function
        End If
        If Not f.NouveauDemande Then Unload f: Exit Function
        Unload f
    End If
    ' pas de specialiste reference pour ce type, ou "autre correspondant"
    Set ChoisirDestinataire = modPatient.ChoisirCorrespondant("", titre)
End Function

' Specialistes actifs pour un ou plusieurs codes TypeExamen (virgules),
' tries par priorite croissante, sous la forme attendue par modCourrier
' (Titre, Prenom, Nom, Specialite, Adresse1/2, CP, Ville, FormuleAppel,
' FormulePolitesse, Tutoiement, BlocDestinataire, ID).
Public Function CorrespondantsPourTypes(ByVal typesExamen As String) As Collection
    Dim res As Collection, fichier As String, specialistes As Collection, parType As Collection
    Dim codes() As String, i As Long, s As Object, l As Object, d As Object
    Dim parId As Object, vus As Object, cle As String
    Set res = New Collection
    fichier = FichierSpecialistes()
    If Len(fichier) = 0 Then Set CorrespondantsPourTypes = res: Exit Function

    Set specialistes = modBase.LireTable(fichier, modConfig.Config("DERIVEES", "FeuilleSpecialistes", "Specialistes"), "ID")
    Set parType = modBase.LireTable(fichier, modConfig.Config("DERIVEES", "FeuilleParType", "Specialistes_ParType"), "ID_Ligne")
    Set parId = CreateObject("Scripting.Dictionary")
    parId.CompareMode = 1
    For Each s In specialistes
        If Not parId.Exists(s("ID")) Then Set parId(s("ID")) = s
    Next s

    Set vus = CreateObject("Scripting.Dictionary")
    vus.CompareMode = 1
    codes = Split(typesExamen, ",")
    For i = LBound(codes) To UBound(codes)
        For Each l In parType
            If StrComp(Trim$(l("TypeExamen")), Trim$(codes(i)), vbTextCompare) = 0 And EstOui(l("Actif")) Then
                cle = l("ID_Specialiste") & "|" & l("NomDestinataire")
                If Not vus.Exists(cle) Then
                    vus(cle) = 1
                    If parId.Exists(l("ID_Specialiste")) Then
                        Set s = parId(l("ID_Specialiste"))
                    Else
                        Set s = Nothing
                    End If
                    Set d = CorrespondantDepuisSpecialiste(s, l)
                    InsererParPriorite res, d
                End If
            End If
        Next l
    Next i
    Set CorrespondantsPourTypes = res
End Function

Private Function FichierSpecialistes() As String
    Dim rel As String, chemin As String
    rel = modConfig.Config("DERIVEES", "FichierSpecialistes", "Base\Correspondants_Specialistes.xlsx")
    If Len(rel) = 0 Then Exit Function
    If InStr(rel, ":") > 0 Or Left$(rel, 2) = "\\" Then chemin = rel Else chemin = modConfig.Racine() & "\" & rel
    If modFichiers.FichierExiste(chemin) Then
        FichierSpecialistes = chemin
    Else
        modLog.LogInfo "Classeur des specialistes absent : " & chemin & " (liste generale utilisee)"
    End If
End Function

Private Function EstOui(ByVal v As String) As Boolean
    v = LCase$(Trim$(v))
    EstOui = (v = "oui" Or v = "1" Or v = "vrai" Or v = "true" Or Len(v) = 0)
End Function

' Fiche correspondant au format modCourrier, a partir d'une ligne
' Specialistes (s, peut etre Nothing) et de sa ligne ParType (l).
Private Function CorrespondantDepuisSpecialiste(ByVal s As Object, ByVal l As Object) As Object
    Dim d As Object, tu As String, prenom As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = "SPT:" & l("ID_Ligne")
    d("NomDestinataire") = l("NomDestinataire")
    tu = LCase$(Trim$(Champ(l, "TutoiementVouvoiement")))
    If Len(tu) = 0 And Not s Is Nothing Then tu = LCase$(Trim$(Champ(s, "TutoiementVouvoiement")))
    If tu <> "tu" Then tu = "vous"
    d("Tutoiement") = tu
    If Not s Is Nothing Then
        d("Titre") = Champ(s, "Titre")
        prenom = Champ(s, "PrenomOuInitiale")
        If prenom = "." Then prenom = ""
        d("Prenom") = prenom
        d("Nom") = Champ(s, "Nom")
        d("Specialite") = Champ(s, "Structure")
        d("Adresse1") = Champ(s, "Adresse1")
        d("Adresse2") = Champ(s, "Adresse2")
        d("CP") = Champ(s, "CodePostal")
        d("Ville") = Champ(s, "Ville")
        d("Tel") = Champ(s, "Telephone")
        d("Structure") = Champ(s, "Structure")
        d("BlocDestinataire") = Replace(Champ(s, "BlocDestinataireComplet"), vbLf, vbCr)
    Else
        d("Titre") = "": d("Prenom") = "": d("Nom") = l("NomDestinataire")
        d("Specialite") = Champ(l, "Structure"): d("Adresse1") = "": d("Adresse2") = ""
        d("CP") = Champ(l, "CodePostal"): d("Ville") = Champ(l, "Ville"): d("Tel") = ""
        d("Structure") = Champ(l, "Structure")
        d("BlocDestinataire") = Replace(Champ(l, "BlocDestinataireComplet"), vbLf, vbCr)
    End If
    If Len(d("BlocDestinataire")) = 0 Then d("BlocDestinataire") = Replace(Champ(l, "BlocDestinataireComplet"), vbLf, vbCr)
    d("Email") = ""
    ' formules : celles de la ligne ParType, puis de la fiche, puis defauts du cabinet
    d("FormuleAppel") = Champ(l, "FormuleAppel")
    If Len(d("FormuleAppel")) = 0 And Not s Is Nothing Then d("FormuleAppel") = Champ(s, "FormuleAppel")
    If Len(d("FormuleAppel")) = 0 Then d("FormuleAppel") = modCourrier.AppelParDefaut(tu = "tu")
    d("FormulePolitesse") = Champ(l, "FormulePolitesse")
    If Len(d("FormulePolitesse")) = 0 Then d("FormulePolitesse") = modCourrier.PolitesseParDefaut(tu = "tu")
    d("Priorite") = Val(Champ(l, "Priorite"))
    d("Actif") = "1"
    Set CorrespondantDepuisSpecialiste = d
End Function

Private Function Champ(ByVal dict As Object, ByVal nom As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(nom) Then Champ = Trim$(CStr(dict(nom)))
End Function

Private Sub InsererParPriorite(ByVal col As Collection, ByVal d As Object)
    Dim i As Long
    For i = 1 To col.Count
        If col(i)("Priorite") > d("Priorite") Then
            col.Add d, , i
            Exit Sub
        End If
    Next i
    col.Add d
End Sub

' ---------------------------------------------------------------------
' Prompt systeme de la demande : gabarit derivee.txt + consignes du type
' ---------------------------------------------------------------------
Public Function ConstruirePrompt(ByVal code As String, ByVal modalite As String, _
                                 ByVal tutoiement As Boolean, ByVal identitePatient As String, _
                                 ByVal motif As String) As String
    Dim d As Object, systeme As String, consigne As String, txtMod As String
    Set d = DefinitionDemande(code)
    If d Is Nothing Then Err.Raise vbObjectError + 503, "modDerivees", "Type de demande inconnu : " & code
    txtMod = TexteModalite(code, modalite)
    consigne = d("ConsigneType")
    consigne = Replace(consigne, "<modalite>", txtMod)
    consigne = Replace(consigne, "<patient>", identitePatient)
    If tutoiement Then
        consigne = Replace(consigne, "<te/vous>", "te")
        consigne = Replace(consigne, "<ton/votre>", "ton")
    Else
        consigne = Replace(consigne, "<te/vous>", "vous")
        consigne = Replace(consigne, "<ton/votre>", "votre")
    End If
    ' les lettres de consultation ne servent PAS de reference de style :
    ' c'est precisement ce qui produisait des demandes en forme de resume
    systeme = modClaude.ChargerPrompt("derivee.txt", False)
    systeme = Replace(systeme, "{{TYPE_DEMANDE}}", LibelleDemande(code, modalite))
    systeme = Replace(systeme, "{{CONSIGNES_TYPE}}", consigne)
    systeme = Replace(systeme, "{{PATIENT}}", identitePatient)
    systeme = Replace(systeme, "{{TUTOIEMENT}}", IIf(tutoiement, _
        "Le destinataire est un confrère proche : tutoiement (""Je te serais reconnaissant"", ""ton service"").", _
        "Le destinataire est vouvoyé (""Je vous serais reconnaissant"", ""votre service"")."))
    If Len(Trim$(motif)) > 0 Then
        systeme = Replace(systeme, "{{MOTIF}}", "Motif ou question précisé par le medecin (a reprendre tel quel, en une phrase) : " & Trim$(motif))
    Else
        systeme = Replace(systeme, "{{MOTIF}}", "Aucun motif particulier n'a été precise : déduis-le sobrement du courrier source, sans le surinterpréter.")
    End If
    ConstruirePrompt = systeme
End Function

' Generation sans interface (testable)
Public Function GenererDerivee(ByVal docSource As Document, ByVal code As String, _
                               ByVal destNouveau As Object, Optional ByVal modalite As String = "", _
                               Optional ByVal motif As String = "") As Document
    Dim pat As Object, corSource As Object, ctx As Object
    Dim corps As String, anonyme As String, problemes As String
    Dim systeme As String, reponse As String, final As String
    Dim nouveauDoc As Document, prog As ufProgression
    Dim identite As String, tutoiement As Boolean, libelle As String

    Set pat = modClaude.PatientDuDocument(docSource)
    Set corSource = modClaude.CorrespondantDuDocument(docSource)
    corps = modCourrier.RecupererCorps(docSource)
    If Len(Trim$(corps)) < 10 Then
        Err.Raise vbObjectError + 500, "modDerivees", "Le courrier source est vide."
    End If
    libelle = LibelleDemande(code, modalite)

    Set ctx = modAnonymise.Construire(pat, corSource)
    ' le nouveau destinataire peut aussi etre cite dans le courrier source
    If Not destNouveau Is Nothing Then
        modAnonymise.AjouterCorrespondant ctx, destNouveau, "DEST2"
    End If
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    ' identite du patient (civilite, prenom, nom, age) : balisee elle aussi
    identite = modAnonymise.Anonymiser(modCourrier.TexteIdentitePatient(pat), ctx)
    problemes = modAnonymise.ScanResiduel(anonyme & vbCr & identite, ctx)
    If Len(problemes) > 0 Then
        If MsgBox("L'anonymisation a detecte un risque avant envoi :" & vbCrLf & vbCrLf & _
                  problemes & vbCrLf & "Envoyer QUAND MEME a l'API ?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Cabinet - protection des donnees") <> vbYes Then
            Exit Function
        End If
    End If

    tutoiement = modCourrier.EstTutoye(destNouveau)
    systeme = ConstruirePrompt(code, modalite, tutoiement, identite, motif)

    Set prog = New ufProgression
    prog.Show vbModeless
    prog.Definir "Redaction de la demande : " & libelle & "..."
    On Error GoTo ErreurApi
    reponse = modClaude.AppelerClaude(systeme, "Patient : " & identite & vbLf & vbLf & _
                                      "Courrier de consultation source :" & vbLf & anonyme)
    On Error GoTo 0
    Unload prog
    Set prog = Nothing

    problemes = modAnonymise.VerifierBalisesRetour(reponse, ctx)
    If Len(problemes) > 0 Then
        modFichiers.EcrireTexteUTF8 modConfig.Chemin("Logs") & "\reponse_rejetee.txt", reponse
        Err.Raise vbObjectError + 501, "modDerivees", _
            "La reponse de l'API a altere des balises d'identite :" & vbCrLf & problemes
    End If
    final = modAnonymise.Reinjecter(reponse, ctx)

    Set nouveauDoc = modCourrier.CreerCourrierPour(pat, destNouveau, "demande - " & libelle)
    modCourrier.RemplacerCorps nouveauDoc, modClaude.NettoyerReponse(final)
    On Error Resume Next
    modGras.AppliquerGras nouveauDoc
    If Err.Number <> 0 Then modLog.LogErreur "Gras lettre derivee : " & Err.Description
    On Error GoTo 0
    Application.StatusBar = "Demande '" & libelle & "' generee : relisez avant validation."
    Set GenererDerivee = nouveauDoc
    Exit Function
ErreurApi:
    If Not prog Is Nothing Then Unload prog
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
