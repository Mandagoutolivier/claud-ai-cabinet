Attribute VB_Name = "modDerivees"
Option Explicit
' =====================================================================
' modDerivees - Lettres derivees du courrier dicte : demande d'examen
' complementaire (echo, Holter, MAPA...) ou d'avis specialise, adressee
' a un confrere choisi, generee par l'API a partir du corps ANONYMISE.
' Le nouveau document reprend l'en-tete du modele (identites posees par
' modCourrier) : l'API ne redige QUE le corps.
' =====================================================================

' Commande principale (Ctrl+Alt+D / commande vocale)
Public Sub LettreDerivee()
    On Error GoTo Erreur
    Dim doc As Document, pat As Object, typeDemande As String, dest As Object

    Set doc = ActiveDocument
    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient (utilisez 'Nouveau courrier').", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If

    typeDemande = ChoisirTypeDemande()
    If Len(typeDemande) = 0 Then Exit Sub

    Set dest = modPatient.ChoisirCorrespondant("", "Destinataire de la demande : " & typeDemande)
    If dest Is Nothing Then Exit Sub

    GenererDerivee doc, typeDemande, dest
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    modLog.LogErreur "LettreDerivee : " & descErr
    MsgBox "Erreur : " & descErr, vbCritical, "Cabinet"
End Sub

Private Function ChoisirTypeDemande() As String
    Dim types() As String, i As Long, items As Collection, d As Object, f As ufListe
    types = Split(modConfig.Config("DERIVEES", "Types", _
        "Échocardiographie transthoracique;Holter ECG;MAPA;Avis spécialisé"), ";")
    Set items = New Collection
    For i = LBound(types) To UBound(types)
        If Len(Trim$(types(i))) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            d("ID") = CStr(i)
            d("Type") = Trim$(types(i))
            items.Add d
        End If
    Next i
    Set f = New ufListe
    f.Configurer "Type de demande", items, Array("Type"), "280 pt"
    f.Show vbModal
    If Not f.Annule Then ChoisirTypeDemande = f.Resultat("Type")
    Unload f
End Function

' Generation sans interface (testable)
Public Function GenererDerivee(ByVal docSource As Document, ByVal typeDemande As String, _
                               ByVal destNouveau As Object) As Document
    Dim pat As Object, corSource As Object, ctx As Object
    Dim corps As String, anonyme As String, problemes As String
    Dim systeme As String, reponse As String, final As String
    Dim nouveauDoc As Document, prog As ufProgression

    Set pat = modClaude.PatientDuDocument(docSource)
    Set corSource = modClaude.CorrespondantDuDocument(docSource)
    corps = modCourrier.RecupererCorps(docSource)
    If Len(Trim$(corps)) < 10 Then
        Err.Raise vbObjectError + 500, "modDerivees", "Le courrier source est vide."
    End If

    Set ctx = modAnonymise.Construire(pat, corSource)
    ' le nouveau destinataire peut aussi etre cite dans le courrier source
    If Not destNouveau Is Nothing Then
        modAnonymise.AjouterCorrespondant ctx, destNouveau, "DEST2"
    End If
    anonyme = modAnonymise.Anonymiser(corps, ctx)
    problemes = modAnonymise.ScanResiduel(anonyme, ctx)
    If Len(problemes) > 0 Then
        If MsgBox("L'anonymisation a detecte un risque avant envoi :" & vbCrLf & vbCrLf & _
                  problemes & vbCrLf & "Envoyer QUAND MEME a l'API ?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Cabinet - protection des donnees") <> vbYes Then
            Exit Function
        End If
    End If

    systeme = modClaude.ChargerPrompt("derivee.txt")
    systeme = Replace(systeme, "{{TYPE_DEMANDE}}", typeDemande)

    Set prog = New ufProgression
    prog.Show vbModeless
    prog.Definir "Redaction de la demande : " & typeDemande & "..."
    On Error GoTo ErreurApi
    reponse = modClaude.AppelerClaude(systeme, anonyme)
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

    Set nouveauDoc = modCourrier.CreerCourrierPour(pat, destNouveau, "demande - " & typeDemande)
    modCourrier.RemplacerCorps nouveauDoc, modClaude.NettoyerReponse(final)
    On Error Resume Next
    modGras.AppliquerGras nouveauDoc
    If Err.Number <> 0 Then modLog.LogErreur "Gras lettre derivee : " & Err.Description
    On Error GoTo 0
    Application.StatusBar = "Demande '" & typeDemande & "' generee : relisez avant validation."
    Set GenererDerivee = nouveauDoc
    Exit Function
ErreurApi:
    If Not prog Is Nothing Then Unload prog
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
