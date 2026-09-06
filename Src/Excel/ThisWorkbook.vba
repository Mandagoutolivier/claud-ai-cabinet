Option Explicit
' Cabinet.xlsm - poste secretaire

Private Sub Workbook_Open()
    On Error Resume Next
    modEchange.DemarrerScrutation
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modEchange.ArreterScrutation
    modAgendaVue.ArreterRafraichissement
End Sub

' On quitte la feuille Agenda : plus de rafraichissement en arriere-plan
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh.Name = "Agenda" Then modAgendaVue.ArreterRafraichissement
End Sub

' Double-clic dans la grille de l'agenda : prise de RDV / actions sur un RDV
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh.Name = "Agenda" Then
        Cancel = True
        modAgendaVue.ClicCreneau Target
    ElseIf Sh.Name = "Honoraires" Then
        Cancel = True
        modHonoraires.ClicLigne Target
    End If
End Sub
