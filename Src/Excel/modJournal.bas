Attribute VB_Name = "modJournal"
Option Explicit
' =====================================================================
' modJournal - Journal des honoraires (Actes\Journal_AAAA.xlsx, feuille
' JOURNAL, une ligne par acte). Exploitable par le comptable.
'
' Trois notions DISTINCTES, jamais deduites l'une de l'autre :
'   - MontantDu        : ce qui est facture au titre de l'acte ;
'   - MontantRegle     : ce qui a ETE RECU (0 tant que rien n'est encaisse) ;
'   - FeuilleSoinsEtat : ou en est la feuille de soins papier (voir ci-dessous).
' Une feuille de soins preparee ou imprimee ne cree aucun encaissement.
'
' Etats de la feuille de soins : (vide) = aucune, Demandee, Imprimee,
' EchecImpression, Remise, Annulee.
' =====================================================================

Public Function EntetesJournal() As Variant
    EntetesJournal = Array("DateActe", "DateSaisie", "SeanceID", "ConsultationID", _
                           "PatientID", "Nom", "Prenom", "DDN", "NIR", _
                           "CodeActe", "CodeOfficiel", "MontantDu", "MontantRegle", _
                           "ModePaiement", "TiersPayant", "Paye", "DateEncaissement", _
                           "FeuilleSoinsEtat", "DateFeuilleSoins", "Notes", "RdvID", "RelanceLe", "ChequeRemisLe")
End Function

Public Sub AssurerJournalAnnee(Optional ByVal annee As Long = 0)
    modBaseIO.CreerClasseurSiAbsent modConfig.FichierJournal(annee), "JOURNAL", EntetesJournal()
    ' journaux crees par une version anterieure : completer les colonnes
    modBaseIO.AssurerColonnes modConfig.FichierJournal(annee), "JOURNAL", EntetesJournal()
End Sub

Public Sub Ajouter(ByVal lignes As Collection, Optional ByVal annee As Long = 0)
    AssurerJournalAnnee annee
    modBaseIO.AjouterLignes modConfig.FichierJournal(annee), "JOURNAL", lignes
End Sub

' Toutes les lignes d'une seance (annee de la date de l'acte)
Public Function LignesSeance(ByVal seanceID As String, Optional ByVal annee As Long = 0) As Collection
    Dim res As Collection, l As Object
    Set res = New Collection
    If Len(seanceID) = 0 Then Set LignesSeance = res: Exit Function
    AssurerJournalAnnee annee
    On Error Resume Next
    For Each l In modBaseIO.LireTableX(modConfig.FichierJournal(annee), "JOURNAL", "SeanceID")
        If l("SeanceID") = seanceID Then res.Add l
    Next l
    On Error GoTo 0
    Set LignesSeance = res
End Function

Public Function SeanceExiste(ByVal seanceID As String, Optional ByVal annee As Long = 0) As Boolean
    SeanceExiste = (LignesSeance(seanceID, annee).Count > 0)
End Function

' Met a jour toutes les lignes d'une seance (etat de feuille de soins,
' reglement...). Renvoie le nombre de lignes modifiees.
Public Function MettreAJourSeance(ByVal seanceID As String, ByVal valeurs As Object, _
                                  Optional ByVal annee As Long = 0) As Long
    AssurerJournalAnnee annee
    MettreAJourSeance = modBaseIO.ModifierLignes(modConfig.FichierJournal(annee), "JOURNAL", _
                                                 "SeanceID", seanceID, valeurs)
End Function

Public Sub OuvrirJournal()
    AssurerJournalAnnee
    On Error GoTo Occupe
    Dim wb As Workbook
    Set wb = Workbooks.Open(modConfig.FichierJournal(), ReadOnly:=True)
    wb.Windows(1).Visible = True     ' fichiers enregistres fenetre masquee (anciennes versions)
    wb.Activate
    Application.StatusBar = "Journal ouvert en lecture seule : les actes s'y ajoutent automatiquement."
    Exit Sub
Occupe:
    MsgBox "Impossible d'ouvrir le journal : " & Err.Description, vbExclamation, "Cabinet"
End Sub
