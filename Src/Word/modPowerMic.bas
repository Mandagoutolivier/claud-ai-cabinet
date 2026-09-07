Attribute VB_Name = "modPowerMic"
Option Explicit
' =====================================================================
' modPowerMic - Points d'entree des boutons du PowerMic (Dragon).
' Dragon execute une commande VBA du type :
'     wd.Run "Cabinet_A_NouvelleLettre"
' Les noms sont propres a Cabinet.dotm (l'ancien complement garde les
' siens, PowerMic_A_NouvelleLettre...) : aucune ambiguite entre les deux.
'
'   A : nouveau courrier pour le patient ARRIVE (medecin traitant de la
'       fiche) ; curseur dans le bloc adresse s'il est vide, sinon dans
'       le corps. Zero clavier.
'   B : formule d'appel selectionnee -> la dictee la remplace.
'   D : finaliser = corriger + lettres de demande a la suite + enregistrer
'       (dossier patient, sortiedragon) + secretariat.
'   P : identite du patient au curseur (equivalent F6).
' =====================================================================

Public Sub Cabinet_A_NouvelleLettre()
    On Error GoTo Erreur
    Dim doc As Document, rng As Range
    modCourrier.NouveauCourrier
    Set doc = ActiveDocument
    ' apres creation : bloc adresse vide -> on s'y place pour dicter le
    ' medecin traitant ; sinon directement dans le corps
    If doc.Bookmarks.Exists("DESTINATAIRE") Then
        If Len(Trim$(doc.Bookmarks("DESTINATAIRE").Range.Text)) = 0 Then
            modCourrier.AllerDestinataire
            Exit Sub
        End If
    End If
    modCourrier.PlacerCurseurCorps doc
    Exit Sub
Erreur:
    MsgBox "Nouvelle lettre impossible : " & Err.Description, vbExclamation, "Cabinet"
End Sub

Public Sub Cabinet_B_FormuleAppel()
    modCourrier.AllerAppel
End Sub

Public Sub Cabinet_D_Finaliser()
    modValidation.FinaliserCourrier
End Sub

Public Sub Cabinet_P_Patient()
    modCourrier.InsererPatient
End Sub

Public Sub Cabinet_Destinataire()
    modCourrier.AllerDestinataire
End Sub

Public Sub Cabinet_Corps()
    modCourrier.AllerCorps
End Sub
