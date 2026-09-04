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
    Set doc = ActiveDocument
    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then
        MsgBox "Ce document n'est pas rattache a un patient (utilisez 'Nouveau courrier').", _
               vbExclamation, "Cabinet"
        Exit Sub
    End If
    typeCourrier = ValiderDocument(doc, False)

    ' Lettres de demande d'examen ou d'avis reperees dans le courrier
    ' principal : generees et transmises sans intervention (modDemandes).
    ' Un courrier qui est lui-meme une demande n'en engendre pas d'autres.
    If Left$(LCase$(typeCourrier), 7) <> "demande" Then
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
    Dim d As Object, typeCourrier As String

    Set pat = modClaude.PatientDuDocument(doc)
    If pat Is Nothing Then Err.Raise vbObjectError + 520, "modValidation", "Document sans patient rattache."
    Set cor = modClaude.CorrespondantDuDocument(doc)

    On Error Resume Next
    typeCourrier = doc.Variables("TypeCourrier")
    On Error GoTo 0
    If Len(typeCourrier) = 0 Then typeCourrier = "courrier"

    dossier = modPatient.DossierPatient(pat)
    base = Format$(Now, "yyyymmdd-hhnn") & " " & modFichiers.NomFichierSur(typeCourrier)
    cheminDocx = dossier & "\" & base & ".docx"
    cheminPdf = dossier & "\" & base & ".pdf"

    doc.SaveAs2 cheminDocx, 12                      ' wdFormatXMLDocument
    doc.ExportAsFixedFormat cheminPdf, 17           ' wdExportFormatPDF

    Set d = CreateObject("Scripting.Dictionary")
    d("PatientID") = pat("ID")
    d("Nom") = pat("Nom")
    d("Prenom") = pat("Prenom")
    d("DDN") = pat("DDN")
    d("NIR") = pat("NIR")
    d("TypeCourrier") = typeCourrier
    If Not cor Is Nothing Then d("DestinataireID") = cor("ID")
    d("CheminDocx") = cheminDocx
    d("CheminPdf") = cheminPdf
    d("DateValidation") = Format$(Now, "dd/mm/yyyy hh:nn")
    d("Poste") = Environ$("COMPUTERNAME")
    modFichiers.EcrireDrapeau modConfig.Chemin("Echange") & "\AEnvoyer", _
                              modFichiers.IdUnique() & "_" & pat("ID"), d

    modLog.LogInfo "Courrier valide : " & cheminDocx
    ValiderDocument = typeCourrier
End Function
