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

' Dossier du patient dans Patients\ (cree si besoin)
Public Function DossierPatient(ByVal pat As Object) As String
    Dim chemin As String
    chemin = modConfig.Chemin("Patients") & "\" & _
             modFichiers.NomFichierSur(pat("Nom") & "_" & pat("Prenom") & "_" & pat("ID"))
    modFichiers.EnsureDossier chemin
    DossierPatient = chemin
End Function
