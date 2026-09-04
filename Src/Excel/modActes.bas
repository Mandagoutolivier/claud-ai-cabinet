Attribute VB_Name = "modActes"
Option Explicit
' =====================================================================
' modActes - Nomenclature des actes (Config\Nomenclature.xlsx, editable
' par la secretaire) et enregistrement d'une seance au journal.
' =====================================================================

Public Function Nomenclature() As Collection
    Dim tous As Collection, actifs As Collection, a As Object
    Set tous = modBaseIO.LireTableX(modConfig.FichierNomenclature(), "ACTES", "Code")
    Set actifs = New Collection
    For Each a In tous
        If a("Actif") <> "0" Then actifs.Add a
    Next a
    Set Nomenclature = actifs
End Function

Public Function ActeParCode(ByVal code As String) As Object
    Dim a As Object
    For Each a In Nomenclature()
        If a("Code") = code Then Set ActeParCode = a: Exit Function
    Next a
    Set ActeParCode = Nothing
End Function

' Enregistre une seance au journal comptable : une ligne par acte (les
' actes associes - ex ECG avec la consultation - generent leur ligne).
' infos : dictionnaire du drapeau (PatientID, Nom, Prenom, DDN, NIR).
' Renvoie le SeanceID.
Public Function EnregistrerSeance(ByVal infos As Object, ByVal actesChoisis As Collection, _
                                  ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                                  ByVal fdsImprimee As Boolean) As String
    Dim seanceID As String, lignes As Collection, a As Object
    seanceID = modFichiers.IdUnique()
    Set lignes = New Collection
    For Each a In actesChoisis
        lignes.Add LigneJournal(seanceID, infos, a("Code"), a("Tarif"), modePaiement, tiersPayant, fdsImprimee)
        If Len(a("CodeAssocie")) > 0 Then
            lignes.Add LigneJournal(seanceID, infos, a("CodeAssocie"), a("TarifAssocie"), _
                                    modePaiement, tiersPayant, fdsImprimee)
        End If
    Next a
    modJournal.Ajouter lignes
    EnregistrerSeance = seanceID
End Function

Private Function LigneJournal(ByVal seanceID As String, ByVal infos As Object, _
                              ByVal code As String, ByVal tarif As String, _
                              ByVal modePaiement As String, ByVal tiersPayant As Boolean, _
                              ByVal fdsImprimee As Boolean) As Object
    Dim d As Object, paye As Boolean
    paye = (LCase$(modePaiement) <> "impaye" And LCase$(modePaiement) <> "impayé")
    Set d = CreateObject("Scripting.Dictionary")
    d("Date") = Format$(Date, "dd/mm/yyyy")
    d("SeanceID") = seanceID
    d("PatientID") = ValeurOuVide(infos, "PatientID")
    d("Nom") = ValeurOuVide(infos, "Nom")
    d("Prenom") = ValeurOuVide(infos, "Prenom")
    d("DDN") = ValeurOuVide(infos, "DDN")
    d("NIR") = ValeurOuVide(infos, "NIR")
    d("CodeActe") = code
    d("Montant") = Replace(tarif, ",", ".")
    d("ModePaiement") = modePaiement
    d("TiersPayant") = IIf(tiersPayant, "O", "N")
    d("Paye") = IIf(paye, "O", "N")
    d("DateEncaissement") = IIf(paye, Format$(Date, "dd/mm/yyyy"), "")
    d("FeuilleSoinsImprimee") = IIf(fdsImprimee, "O", "N")
    Set LigneJournal = d
End Function

Private Function ValeurOuVide(ByVal dict As Object, ByVal cle As String) As String
    If dict.Exists(cle) Then ValeurOuVide = dict(cle) Else ValeurOuVide = ""
End Function

' Lignes pour la feuille de soins : un element par acte, associes compris
Public Function LignesPourImpression(ByVal actesChoisis As Collection) As Collection
    Dim col As New Collection, a As Object, d As Object
    For Each a In actesChoisis
        Set d = CreateObject("Scripting.Dictionary")
        d("CodeActe") = ChoixLibelleCerfa(a, a("Code"))
        d("Montant") = a("Tarif")
        col.Add d
        If Len(a("CodeAssocie")) > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d("CodeActe") = a("CodeAssocie")
            d("Montant") = a("TarifAssocie")
            col.Add d
        End If
    Next a
    Set LignesPourImpression = col
End Function

Private Function ChoixLibelleCerfa(ByVal a As Object, ByVal defaut As String) As String
    If a.Exists("LibelleCerfa") Then
        If Len(a("LibelleCerfa")) > 0 Then
            ChoixLibelleCerfa = a("LibelleCerfa")
            Exit Function
        End If
    End If
    ChoixLibelleCerfa = defaut
End Function

' Total d'une selection d'actes (associes compris)
Public Function TotalActes(ByVal actesChoisis As Collection) As Double
    Dim a As Object, total As Double
    For Each a In actesChoisis
        total = total + Val(Replace(a("Tarif"), ",", "."))
        If Len(a("CodeAssocie")) > 0 Then total = total + Val(Replace(a("TarifAssocie"), ",", "."))
    Next a
    TotalActes = total
End Function
