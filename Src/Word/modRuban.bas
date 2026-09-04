Attribute VB_Name = "modRuban"
Option Explicit
' =====================================================================
' modRuban - Rappels du ruban "Cabinet" (customUI14.xml injecte dans
' Cabinet.dotm par build.ps1). Un seul point d'entree : Ruban_Action.
' Les macros sont appelees par leur nom (Application.Run) : ce sont les
' memes que celles des raccourcis clavier et des commandes Dragon.
' =====================================================================

Public Sub Ruban_Action(ByVal control As IRibbonControl)
    On Error GoTo Erreur
    Dim macro As String
    Select Case control.ID
        Case "cabNouveau":      macro = "NouveauCourrier"
        Case "cabPatient":      macro = "InsererPatient"
        Case "cabCorriger":     macro = "CorrigerCourrier"
        Case "cabDerivee":      macro = "LettreDerivee"
        Case "cabValider":      macro = "ValiderCourrier"
        Case "cabEcg":          macro = "EnvoyerECG"
        Case "cabGras":         macro = "MettreEnGras"
        Case "cabMedicaments":  macro = "OuvrirDictionnaireMedicaments"
        Case "cabExpressions":  macro = "OuvrirDictionnaireExpressions"
        Case "cabAide":         macro = "AideCabinet"
        Case "cabDiagnostic":   macro = "DiagnosticCabinet"
        Case Else
            MsgBox "Bouton inconnu : " & control.ID, vbExclamation, "Cabinet"
            Exit Sub
    End Select
    Application.Run macro
    Exit Sub
Erreur:
    modLog.LogErreur "Ruban " & control.ID & " : " & Err.Description
    MsgBox Err.Description, vbCritical, "Cabinet"
End Sub
