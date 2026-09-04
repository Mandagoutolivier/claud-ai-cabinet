Attribute VB_Name = "modTexte"
Option Explicit
' =====================================================================
' modTexte - Utilitaires texte : pliage (minuscules sans accents) pour
' les recherches et l'anonymisation insensibles a la casse/aux accents.
' =====================================================================

Public Function Plier(ByVal s As String) As String
    Dim src As String, dst As String, i As Long, p As Long, c As String, res As String
    s = LCase$(s)
    ' digrammes d'abord
    s = Replace(s, ChrW$(230), "ae")       ' ae lie
    s = Replace(s, ChrW$(339), "oe")       ' oe lie
    src = ChrW$(224) & ChrW$(226) & ChrW$(228) & ChrW$(225) & ChrW$(227) & ChrW$(229) & _
          ChrW$(233) & ChrW$(232) & ChrW$(234) & ChrW$(235) & _
          ChrW$(237) & ChrW$(236) & ChrW$(238) & ChrW$(239) & _
          ChrW$(243) & ChrW$(242) & ChrW$(244) & ChrW$(246) & ChrW$(245) & _
          ChrW$(250) & ChrW$(249) & ChrW$(251) & ChrW$(252) & _
          ChrW$(253) & ChrW$(255) & ChrW$(231) & ChrW$(241)
    dst = "aaaaaaeeeeiiiiooooouuuuyycn"
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        p = InStr(src, c)
        If p > 0 Then res = res & Mid$(dst, p, 1) Else res = res & c
    Next i
    Plier = res
End Function

' Civilite d'apres le sexe (M/F)
Public Function Civilite(ByVal sexe As String) As String
    If UCase$(Left$(sexe, 1)) = "F" Then Civilite = "Madame" Else Civilite = "Monsieur"
End Function

Public Function CiviliteCourte(ByVal sexe As String) As String
    If UCase$(Left$(sexe, 1)) = "F" Then CiviliteCourte = "Mme" Else CiviliteCourte = "M."
End Function

Public Function NeLe(ByVal sexe As String) As String
    If UCase$(Left$(sexe, 1)) = "F" Then NeLe = "née le" Else NeLe = "né le"
End Function
