Option Explicit
' Petite fenetre non bloquante affichee pendant les appels a l'API.

Public Sub Definir(ByVal msg As String)
    lblMsg.Caption = msg
    Me.Repaint
    DoEvents
End Sub
