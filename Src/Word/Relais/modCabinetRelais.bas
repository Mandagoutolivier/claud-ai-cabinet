Attribute VB_Name = "modCabinetRelais"
Option Explicit
' =====================================================================
' modCabinetRelais - RELAIS installe dans Normal.dotm par deploy.ps1.
' Sur ce poste, les raccourcis clavier definis dans un modele global
' (Cabinet.dotm) ne s'executent pas de facon fiable ; ceux definis dans
' Normal.dotm vers des macros de Normal.dotm, si. Ces relais appellent
' donc les vraies macros de Cabinet.dotm par Application.Run.
' Ne contient aucune logique metier : Cabinet.dotm reste la seule source.
' =====================================================================

Private Sub CAB_Relais(ByVal macro As String)
    On Error GoTo Absent
    Application.Run macro
    Exit Sub
Absent:
    MsgBox "Le module Cabinet.dotm n'est pas charge dans Word (" & macro & ")." & vbCrLf & _
           "Verifiez le dossier de demarrage de Word ou relancez deploy.ps1.", vbExclamation, "Cabinet"
End Sub

Public Sub CAB_NouveauCourrier()
    CAB_Relais "NouveauCourrier"
End Sub

Public Sub CAB_CorrigerCourrier()
    CAB_Relais "CorrigerCourrier"
End Sub

Public Sub CAB_LettreDerivee()
    CAB_Relais "LettreDerivee"
End Sub

Public Sub CAB_InsererPatient()
    CAB_Relais "InsererPatient"
End Sub

Public Sub CAB_EnvoyerECG()
    CAB_Relais "EnvoyerECG"
End Sub

Public Sub CAB_ValiderCourrier()
    CAB_Relais "ValiderCourrier"
End Sub

Public Sub CAB_MettreEnGras()
    CAB_Relais "MettreEnGras"
End Sub

Public Sub CAB_Aide()
    CAB_Relais "AideCabinet"
End Sub

' Sonde de test (fichier temoin) : utilisee par installer_relais.ps1
Public Sub CAB_Sonde()
    CAB_Relais "SondeRaccourci"
End Sub
