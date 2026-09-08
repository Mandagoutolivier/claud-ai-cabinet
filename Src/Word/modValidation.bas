Attribute VB_Name = "modValidation"
Option Explicit
' =====================================================================
' modValidation - Validation du courrier par le medecin (Ctrl+Alt+V) :
'  1. enregistre le docx + un PDF dans le dossier du patient
'  2. depose un fichier-drapeau dans Echange\AEnvoyer\ ; le poste
'     secretaire le detecte, imprime, choisit l'acte, alimente le
'     journal comptable et la feuille de soins.
' =====================================================================

Public Sub ValiderCourrier()
    On Error GoTo Erreur
    Dim doc As Document, pat As Object, typeCourrier As String, rapport As String
    Dim memeDoc As Boolean
    Set doc = ActiveDocument
    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient : appuyez sur F6 (ou Ctrl+Alt+P) pour le choisir.", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    typeCourrier = VariableDoc(doc, "TypeCourrier")
    If Len(typeCourrier) = 0 Then typeCourrier = "courrier"
    memeDoc = modConfig.ConfigBool("DERIVEES", "MemeDocument", True)

    ' Lettres de demande d'examen ou d'avis reperees dans le courrier
    ' principal (modDemandes). Un courrier qui est lui-meme une demande n'en
    ' engendre pas d'autres.
    ' MemeDocument : elles sont ajoutees AVANT l'enregistrement, pour que le
    ' fichier unique (dossier patient, sortie, secretariat) les contienne.
    If memeDoc And Left$(LCase$(typeCourrier), 7) <> "demande" Then
        On Error Resume Next
        rapport = modDemandes.GenererDemandesAutomatiques(doc)
        If Err.Number <> 0 Then
            modLog.LogErreur "Demandes automatiques : " & Err.Description
            rapport = "Lettres de demande non generees : " & Err.Description
        End If
        On Error GoTo Erreur
        doc.Activate
    End If

    typeCourrier = ValiderDocument(doc, False)

    If Not memeDoc And Left$(LCase$(typeCourrier), 7) <> "demande" Then
        On Error Resume Next
        rapport = modDemandes.GenererDemandesAutomatiques(doc)
        If Err.Number <> 0 Then
            modLog.LogErreur "Demandes automatiques : " & Err.Description
            rapport = "Lettres de demande non generees : " & Err.Description
        End If
        On Error GoTo Erreur
    End If

    MsgBox "Courrier valide et transmis au secretariat." & vbCrLf & _
           "(" & pat("Prenom") & " " & pat("Nom") & " - " & typeCourrier & ")" & _
           IIf(Len(rapport) > 0, vbCrLf & vbCrLf & "Lettres de demande :" & vbCrLf & rapport, ""), _
           vbInformation, "Cabinet"
    Exit Sub
Erreur:
    Dim descErr As String
    descErr = Err.Description
    modLog.LogErreur "ValiderCourrier : " & descErr
    MsgBox "Erreur lors de la validation : " & descErr, vbCritical, "Cabinet"
End Sub

' Validation d'un document rattache a un patient : docx + PDF dans le
' dossier du patient, drapeau pour le secretariat. Renvoie le type de
' courrier. silencieux=True : aucun message (lettres derivees automatiques).
Public Function ValiderDocument(ByVal doc As Document, ByVal silencieux As Boolean) As String
    Dim pat As Object, cor As Object
    Dim dossier As String, base As String, cheminDocx As String, cheminPdf As String
    Dim d As Object, typeCourrier As String, consultationID As String, dateActe As String

    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then Err.Raise vbObjectError + 520, "modValidation", _
        "Document sans patient rattache : appuyez sur F6 (ou Ctrl+Alt+P) pour choisir le patient, puis validez."
    Set cor = modClaude.CorrespondantDuDocument(doc)
    If cor Is Nothing Then
        ' destinataire dicte (saisie rapide) : reconnu dans la base si possible
        If Len(modCourrier.ReconnaitreDestinataire(doc)) > 0 Then Set cor = modClaude.CorrespondantDuDocument(doc)
        modCourrier.MettreEnFormeDestinataire doc
    End If

    typeCourrier = VariableDoc(doc, "TypeCourrier")
    If Len(typeCourrier) = 0 Then typeCourrier = "courrier"

    ' Identifiant stable du document : attribue a la premiere validation et
    ' conserve dans le document. Une revalidation apres correction produit
    ' une nouvelle VERSION du meme acte, jamais un acte supplementaire.
    consultationID = VariableDoc(doc, "ConsultationID")
    If Len(consultationID) = 0 Then
        consultationID = modFichiers.IdUnique()
        DefinirVariableDoc doc, "ConsultationID", consultationID
    End If

    ' Date REELLE de l'acte : celle de la consultation si elle est portee par
    ' le document, sinon celle de la premiere validation (figee ensuite).
    dateActe = VariableDoc(doc, "DateActe")
    If Not modTexte.DateFrValide(dateActe) Then
        dateActe = Format$(Date, "dd/mm/yyyy")
        DefinirVariableDoc doc, "DateActe", dateActe
    End If

    dossier = modPatient.DossierPatient(pat)
    base = Format$(modTexte.DateFr(dateActe), "yyyy-mm-dd") & " " & _
           modFichiers.NomFichierSur(typeCourrier) & " " & consultationID
    cheminDocx = NomVersionne(dossier, base, "docx")
    cheminPdf = Left$(cheminDocx, Len(cheminDocx) - 4) & "pdf"

    doc.SaveAs2 cheminDocx, 12                      ' wdFormatXMLDocument
    doc.ExportAsFixedFormat cheminPdf, 17           ' wdExportFormatPDF

    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = modTexte.DdnPatient(pat)
    d("NIR") = pat("NIR")
    d("TypeCourrier") = typeCourrier
    d("ConsultationID") = consultationID
    d("SeanceID") = consultationID
    d("DateActe") = dateActe
    d("RdvID") = VariableDoc(doc, "RdvID")
    If Not cor Is Nothing Then d("DestinataireID") = cor("ID")
    d("CheminDocx") = cheminDocx
    d("CheminPdf") = cheminPdf
    d("DateValidation") = Format$(Now, "dd/mm/yyyy hh:nn")
    d("Poste") = Environ$("COMPUTERNAME")
    modFichiers.EcrireDrapeau modConfig.Chemin("Echange") & "\AEnvoyer", _
                              modFichiers.IdUnique() & "_" & pat("ID"), d

    ' copie dans le dossier de sortie du cabinet ([SORTIE] Dossier, ex :
    ' \\DS224\home\sortiedragon) sous "NOM Prenom aammjjhhmm.docx", comme
    ' l'ancien modele : le secretariat y retrouve le fichier complet.
    CopierVersSortie cheminDocx, pat, typeCourrier

    modLog.LogInfo "Courrier valide : " & cheminDocx & " (consultation " & consultationID & ")"
    ValiderDocument = typeCourrier
End Function

Private Sub CopierVersSortie(ByVal cheminDocx As String, ByVal pat As Object, ByVal typeCourrier As String)
    On Error Resume Next
    Dim dossier As String, dest As String
    dossier = modConfig.Config("SORTIE", "Dossier", "")
    If Len(dossier) = 0 Then Exit Sub
    If Left$(LCase$(typeCourrier), 7) = "demande" Then Exit Sub    ' les demandes sont dans le fichier principal
    modFichiers.EnsureDossier dossier
    dest = dossier & "\" & modFichiers.NomFichierSur(UCase$(pat("Nom")) & " " & pat("Prenom") & " " & Format$(Now, "yymmddhhnn")) & ".docx"
    FileCopy cheminDocx, dest
    If Err.Number <> 0 Then
        modLog.LogErreur "Copie vers la sortie impossible (" & dest & ") : " & Err.Description
    Else
        modLog.LogInfo "Copie de sortie : " & dest
    End If
End Sub

' UNE TOUCHE (bouton D du PowerMic) : corrige le courrier, ajoute les
' lettres de demande a la suite, enregistre (dossier patient + dossier de
' sortie) et transmet au secretariat.
Public Sub FinaliserCourrier()
    On Error GoTo Erreur
    Dim doc As Document
    Set doc = ActiveDocument
    If modClaude.PatientDuDocument(doc) Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient : appuyez sur F6 pour le choisir, puis recommencez.", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    Application.StatusBar = "Finalisation : correction..."
    If Not modClaude.CorrigerDocument(doc, True) Then Exit Sub     ' message deja affiche
    Application.StatusBar = "Finalisation : lettres de demande et enregistrement..."
    ValiderCourrier
    Exit Sub
Erreur:
    MsgBox "Finalisation impossible : " & Err.Description, vbCritical, "Cabinet"
End Sub

' Chemin libre : base.ext, puis "base v2.ext", "base v3.ext"... Les
' corrections successives sont conservees au lieu de s'ecraser.
Private Function NomVersionne(ByVal dossier As String, ByVal base As String, _
                              ByVal ext As String) As String
    Dim chemin As String, n As Long
    chemin = dossier & "\" & base & "." & ext
    n = 1
    Do While Len(Dir$(chemin)) > 0
        n = n + 1
        chemin = dossier & "\" & base & " v" & n & "." & ext
        If n > 200 Then Exit Do
    Loop
    NomVersionne = chemin
End Function

Private Function VariableDoc(ByVal doc As Document, ByVal nom As String) As String
    On Error Resume Next
    VariableDoc = doc.Variables(nom).Value
    Err.Clear
End Function

Private Sub DefinirVariableDoc(ByVal doc As Document, ByVal nom As String, ByVal valeur As String)
    On Error Resume Next
    doc.Variables(nom).Value = valeur
    Err.Clear
End Sub
