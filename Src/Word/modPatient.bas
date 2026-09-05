Attribute VB_Name = "modPatient"
Option Explicit
' =====================================================================
' modPatient - Selection du patient et du correspondant (poste medecin).
' =====================================================================

Public Function ChoisirPatient() As Object
    Dim f As ufListe, patients As Collection
    Set patients = modBase.Patients(True)
    modLog.Etape "selecteur patient : creation du formulaire"
    Set f = New ufListe
    modLog.Etape "selecteur patient : remplissage (" & patients.Count & " patients)"
    f.Configurer "Choisir un patient (tapez pour filtrer)", patients, _
                 Array("Nom", "Prenom", "DDN", "Ville", "ID"), "120 pt;90 pt;70 pt;80 pt;40 pt"
    modLog.Etape "selecteur patient : affichage"
    f.Show vbModal
    If Not f.Annule Then Set ChoisirPatient = f.Resultat
    Unload f
End Function

' preselectionID : correspondant preselectionne (ex : medecin traitant),
' il suffit alors de valider par Entree.
Public Function ChoisirCorrespondant(Optional ByVal preselectionID As String = "", _
                                     Optional ByVal titre As String = "") As Object
    Dim f As ufListe
    If Len(titre) = 0 Then titre = "Choisir le destinataire"
    Set f = New ufListe
    f.Configurer titre, modBase.Correspondants(True), _
                 Array("Titre", "Nom", "Prenom", "Specialite", "Ville"), _
                 "30 pt;110 pt;80 pt;120 pt;80 pt", preselectionID
    f.Show vbModal
    If Not f.Annule Then Set ChoisirCorrespondant = f.Resultat
    Unload f
End Function

' Dossier du patient. Le nom du dossier commence TOUJOURS par l'identifiant
' stable (ID) : un changement d'etat civil renomme le dossier au lieu d'en
' creer un second. Les anciens dossiers "Nom_Prenom_ID" sont migres.
Public Function DossierPatient(ByVal pat As Object) As String
    Dim racine As String, id As String, souhaite As String, existant As String
    racine = modConfig.Chemin("Patients")
    modFichiers.EnsureDossier racine
    id = Trim$(CStr(pat("ID")))
    If Len(id) = 0 Then
        Err.Raise vbObjectError + 630, "modPatient", _
            "Patient sans identifiant : impossible de determiner son dossier."
    End If
    souhaite = racine & "\" & modFichiers.NomFichierSur( _
                   id & "_" & pat("Nom") & "_" & pat("Prenom"))
    existant = DossierExistant(racine, id)
    If Len(existant) > 0 Then
        If StrComp(existant, souhaite, vbTextCompare) <> 0 Then
            On Error Resume Next
            Name existant As souhaite
            If Err.Number <> 0 Then
                ' renommage impossible (dossier ouvert, droits) : on garde l'existant
                Err.Clear
                On Error GoTo 0
                DossierPatient = existant
                Exit Function
            End If
            On Error GoTo 0
        Else
            DossierPatient = existant
            Exit Function
        End If
    End If
    modFichiers.EnsureDossier souhaite
    DossierPatient = souhaite
End Function

' Cherche un dossier deja present pour cet ID : nouveau schema "ID_..."
' ou ancien schema "Nom_Prenom_ID".
Private Function DossierExistant(ByVal racine As String, ByVal id As String) As String
    ' FileSystemObject (et non Dir$) : DossierPatient peut etre appele depuis
    ' une boucle Dir$ de l'appelant, qui serait alors reinitialisee.
    Dim fso As Object, dossier As Object, nom As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(racine) Then Exit Function
    For Each dossier In fso.GetFolder(racine).SubFolders
        nom = dossier.Name
        If LCase$(Left$(nom, Len(id) + 1)) = LCase$(id & "_") Or _
           LCase$(Right$(nom, Len(id) + 1)) = LCase$("_" & id) Then
            DossierExistant = dossier.Path
            Exit Function
        End If
    Next dossier
End Function
