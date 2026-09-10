Attribute VB_Name = "modDemandesR12"
Option Explicit
' =====================================================================
' modDemandesR12 - Architecture R12 des lettres de demande.
'
' Un SEUL appel a l'API, celui de la correction du courrier principal :
' le prompt de correction est complete par les consignes R12 et l'API rend
'   ---CORPS_COURRIER--- ... ---FIN_CORPS_COURRIER---
' suivi de zero, un ou plusieurs blocs
'   ---DEMANDE_DESTINATION---
'   CLE_DESTINATION=CODE_PROFIL   (ou A_COMPLETER)
'   ---CORPS_DESTINATION---
'   corps de la lettre de demande
'   ---FIN_CORPS_DESTINATION---
'   ---FIN_DEMANDE_DESTINATION---
' Les corps des demandes sont memorises dans le document (variables) puis,
' a la finalisation, le VBA assemble chaque lettre depuis le modele du
' cabinet (destinataire de la base des specialistes selon le code du
' profil, appel, corps, politesse) sans nouvel appel a l'API.
' Si le courrier contient une formule habituelle de demande (dictionnaires
' Config\demandes) mais que l'API n'a rendu aucun bloc, un second appel
' "demandes seules" est fait ([DERIVEES] SecondAppel=1).
' Active par [DERIVEES] Mode=R12 (defaut) ; Mode=Profils = un appel API
' par demande d'apres son profil (ancien fonctionnement).
' =====================================================================

Public Const B_CORPS As String = "---CORPS_COURRIER---"
Public Const B_FIN_CORPS As String = "---FIN_CORPS_COURRIER---"
Public Const B_DEMANDE As String = "---DEMANDE_DESTINATION---"
Public Const B_FIN_DEMANDE As String = "---FIN_DEMANDE_DESTINATION---"
Public Const B_CLE As String = "CLE_DESTINATION="
Public Const B_CORPS_DEST As String = "---CORPS_DESTINATION---"
Public Const B_FIN_CORPS_DEST As String = "---FIN_CORPS_DESTINATION---"
Private Const SEP_DEMANDE As String = "|||DEMANDE|||"
Private Const SEP_CHAMP As String = "|||"

Public Function ModeR12() As Boolean
    ModeR12 = (UCase$(modConfig.Config("DERIVEES", "Mode", "R12")) = "R12")
End Function

' --- construction du prompt ---------------------------------------------

' Consignes ajoutees au prompt de correction : texte de
' Config\prompts\demandes_r12.txt + regles de destination (une ligne par
' profil : libelle, code, destinataire connu ou A_COMPLETER) + format.
Public Function ConsignesDemandes() As String
    Dim s As String, chemin As String
    chemin = modConfig.Chemin("Config") & "\prompts\demandes_r12.txt"
    If modFichiers.FichierExiste(chemin) Then
        s = modFichiers.LireTexteUTF8(chemin)
    Else
        s = "Après le bloc du courrier principal, produis un bloc de demande complémentaire pour chaque examen ou avis explicitement décidé dans le courrier."
    End If
    s = Replace(s, "{{REGLES_DESTINATION}}", ReglesDestination())
    s = Replace(s, "{{FORMAT_DEMANDE}}", FormatDemande())
    ConsignesDemandes = vbLf & vbLf & s
End Function

Private Function ReglesDestination() As String
    Dim s As String, profils As Collection, d As Object, p As Object, dest As Object, cle As String, qui As String
    On Error Resume Next
    Set profils = modDemandes.ListerProfils()
    On Error GoTo 0
    If profils Is Nothing Then ReglesDestination = "- Tout examen ou avis : " & B_CLE & "A_COMPLETER": Exit Function
    For Each d In profils
        Set p = Nothing: Set dest = Nothing
        On Error Resume Next
        Set p = modDemandes.ChargerProfil(d("Code"))
        If Not p Is Nothing Then Set dest = modDerivees.DestinataireAutomatique(p)
        On Error GoTo 0
        cle = UCase$(d("Code"))
        qui = ""
        If Not dest Is Nothing Then
            If dest("ID") <> "A_COMPLETER" Then qui = " (" & dest("NomDestinataire") & ")"
        End If
        s = s & "- " & d("Libelle") & " : " & B_CLE & cle & qui & vbLf
    Next d
    s = s & "- Tout autre examen, avis ou orientation, ou correspondant précis non connu : " & B_CLE & "A_COMPLETER" & vbLf
    ReglesDestination = s
End Function

Private Function FormatDemande() As String
    FormatDemande = B_DEMANDE & vbLf & B_CLE & "CLE_EXACTE" & vbLf & B_CORPS_DEST & vbLf & _
                    "Corps médical de la demande" & vbLf & B_FIN_CORPS_DEST & vbLf & B_FIN_DEMANDE
End Function

' --- lecture de la reponse ----------------------------------------------

' Corps du courrier principal : le bloc CORPS_COURRIER s'il existe, sinon la
' reponse privee de ses blocs de demande (l'API a ignore les balises).
Public Function ExtraireCorps(ByVal reponse As String) As String
    Dim a As Long, b As Long
    a = InStr(1, reponse, B_CORPS, vbTextCompare)
    If a > 0 Then
        b = InStr(a + Len(B_CORPS), reponse, B_FIN_CORPS, vbTextCompare)
        If b > 0 Then
            ExtraireCorps = Nettoyer(Mid$(reponse, a + Len(B_CORPS), b - a - Len(B_CORPS)))
            Exit Function
        End If
    End If
    a = InStr(1, reponse, B_DEMANDE, vbTextCompare)
    If a > 0 Then reponse = Left$(reponse, a - 1)
    ExtraireCorps = Nettoyer(Replace(Replace(reponse, B_CORPS, ""), B_FIN_CORPS, ""))
End Function

' Blocs de demande : Collection de dictionnaires (Cle, Corps)
Public Function ExtraireDemandes(ByVal reponse As String) As Collection
    Dim res As Collection, pos As Long, a As Long, b As Long, bloc As String, d As Object
    Dim c1 As Long, c2 As Long, k As Long, fin As Long, cle As String
    Set res = New Collection
    pos = 1
    Do
        a = InStr(pos, reponse, B_DEMANDE, vbTextCompare)
        If a = 0 Then Exit Do
        b = InStr(a + Len(B_DEMANDE), reponse, B_FIN_DEMANDE, vbTextCompare)
        If b = 0 Then Exit Do
        bloc = Mid$(reponse, a, b - a + Len(B_FIN_DEMANDE))
        pos = b + Len(B_FIN_DEMANDE)
        k = InStr(1, bloc, B_CLE, vbTextCompare)
        c1 = InStr(1, bloc, B_CORPS_DEST, vbTextCompare)
        c2 = InStr(1, bloc, B_FIN_CORPS_DEST, vbTextCompare)
        If k > 0 And c1 > 0 And c2 > c1 Then
            fin = InStr(k, bloc, vbLf)
            If fin = 0 Then fin = InStr(k, bloc, vbCr)
            If fin = 0 Then fin = c1
            cle = UCase$(Trim$(Replace(Replace(Mid$(bloc, k + Len(B_CLE), fin - k - Len(B_CLE)), vbCr, ""), vbLf, "")))
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            d("Cle") = cle
            d("Corps") = Nettoyer(Mid$(bloc, c1 + Len(B_CORPS_DEST), c2 - c1 - Len(B_CORPS_DEST)))
            If Len(d("Corps")) > 0 Then res.Add d
        End If
    Loop
    Set ExtraireDemandes = res
End Function

Private Function Nettoyer(ByVal t As String) As String
    t = Replace(t, "**", "")                     ' gras Markdown : modGras s'en charge
    Nettoyer = Trim$(modClaude.NettoyerReponse(t))
End Function

' --- memorisation dans le document --------------------------------------

Public Sub MemoriserDemandes(ByVal doc As Document, ByVal demandes As Collection)
    Dim d As Object, s As String
    For Each d In demandes
        If Len(s) > 0 Then s = s & SEP_DEMANDE
        s = s & d("Cle") & SEP_CHAMP & d("Corps")
    Next d
    doc.Variables("DemandesR12") = IIf(Len(s) > 0, s, " ")
    doc.Variables("DemandesR12Nb") = CStr(demandes.Count)
End Sub

Public Function DemandesMemorisees(ByVal doc As Document) As Collection
    Dim res As Collection, s As String, blocs() As String, i As Long, p As Long, d As Object
    Set res = New Collection
    On Error Resume Next
    s = doc.Variables("DemandesR12")
    On Error GoTo 0
    If Len(Trim$(s)) = 0 Then Set DemandesMemorisees = res: Exit Function
    blocs = Split(s, SEP_DEMANDE)
    For i = LBound(blocs) To UBound(blocs)
        p = InStr(blocs(i), SEP_CHAMP)
        If p > 0 Then
            Set d = CreateObject("Scripting.Dictionary")
            d.CompareMode = 1
            d("Cle") = Left$(blocs(i), p - 1)
            d("Corps") = Mid$(blocs(i), p + Len(SEP_CHAMP))
            res.Add d
        End If
    Next i
    Set DemandesMemorisees = res
End Function

Public Function DemandesDejaAnalysees(ByVal doc As Document) As Boolean
    Dim s As String
    On Error Resume Next
    s = doc.Variables("DemandesR12Nb")
    On Error GoTo 0
    DemandesDejaAnalysees = (Len(s) > 0)
End Function

' --- second appel "demandes seules" -------------------------------------
' Le courrier contient une formule habituelle de demande mais l'API n'a
' rendu aucun bloc : on le lui redemande explicitement.
Public Function DemandesParSecondAppel(ByVal anonyme As String) As Collection
    Dim systeme As String, reponse As String
    systeme = "Tu es l'assistant de rédaction d'un cardiologue libéral français. À partir du courrier de consultation anonymisé fourni (identités masquées par des balises {{...}}, à recopier strictement), produis UNIQUEMENT les blocs de demande d'examen ou d'avis explicitement décidés dans ce courrier, sans bloc CORPS_COURRIER ni commentaire." & _
              ConsignesDemandes()
    reponse = modClaude.AppelerClaude(systeme, anonyme)
    Set DemandesParSecondAppel = ExtraireDemandes(reponse)
End Function

' --- assemblage des lettres (sans API) ----------------------------------

' Destinataire pour une cle : code de profil -> destinataire automatique
' de la base des specialistes ; A_COMPLETER ou inconnu -> bloc a completer.
Public Function DestinatairePourCle(ByVal cle As String, ByRef libelle As String) As Object
    Dim p As Object
    libelle = ""
    If Len(cle) > 0 And cle <> "A_COMPLETER" Then
        On Error Resume Next
        If modDemandes.ProfilExiste(cle) Then Set p = modDemandes.ChargerProfil(cle)
        On Error GoTo 0
    End If
    If p Is Nothing Then
        On Error Resume Next
        Set p = modDemandes.ChargerProfil("AUTRE_EXAMEN")
        On Error GoTo 0
    End If
    If p Is Nothing Then
        libelle = "demande"
        Set DestinatairePourCle = DestinataireACompleter(libelle)
        Exit Function
    End If
    libelle = modDemandes.LibelleProfil(p)
    If cle = "A_COMPLETER" Or Not modDemandes.ProfilExiste(cle) Then
        Set DestinatairePourCle = DestinataireACompleter(libelle)
    Else
        Set DestinatairePourCle = modDerivees.DestinataireAutomatique(p)
    End If
End Function

Private Function DestinataireACompleter(ByVal libelle As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ID") = "A_COMPLETER"
    d("NomDestinataire") = "Destinataire à compléter"
    d("Titre") = "": d("Prenom") = "": d("Nom") = ""
    d("Specialite") = "": d("Adresse1") = "": d("Adresse2") = "": d("CP") = "": d("Ville") = ""
    d("Tel") = "": d("Email") = "": d("Structure") = ""
    d("BlocDestinataire") = "DESTINATAIRE À COMPLÉTER" & vbCr & "(" & libelle & ")"
    d("Tutoiement") = "vous"
    d("FormuleAppel") = modCourrier.AppelParDefaut(False)
    d("FormulePolitesse") = modCourrier.PolitesseParDefaut(False)
    d("Priorite") = 999
    d("Actif") = "1"
    Set DestinataireACompleter = d
End Function

' Lettre de demande assemblee depuis le modele du cabinet : en-tete
' (destinataire, appel, politesse selon la fiche), corps rendu par l'API,
' gras par dictionnaires. Aucun appel reseau.
Public Function AssemblerLettre(ByVal docSource As Document, ByVal pat As Object, ByVal d As Object, _
                                ByRef dest As Object, ByRef libelle As String) As Document
    Dim nouveau As Document
    Set dest = DestinatairePourCle(d("Cle"), libelle)
    Set nouveau = modCourrier.CreerCourrierPour(pat, dest, "demande - " & libelle)
    nouveau.Variables("ProfilDemande") = d("Cle")
    ' les demandes gardent appel et politesse de la fiche meme quand le
    ' medecin dicte les siennes sur le courrier principal
    If nouveau.Bookmarks.Exists("APPEL") Then
        If Len(Trim$(nouveau.Bookmarks("APPEL").Range.Text)) = 0 Then
            modCourrier.RemplirSignet nouveau, "APPEL", IIf(Len(dest("FormuleAppel")) > 0, dest("FormuleAppel"), modCourrier.AppelParDefaut(modCourrier.EstTutoye(dest), dest))
        End If
    End If
    If nouveau.Bookmarks.Exists("POLITESSE") Then
        If Len(Trim$(nouveau.Bookmarks("POLITESSE").Range.Text)) = 0 Then
            modCourrier.RemplirSignet nouveau, "POLITESSE", IIf(Len(dest("FormulePolitesse")) > 0, dest("FormulePolitesse"), modCourrier.PolitesseParDefaut(modCourrier.EstTutoye(dest)))
        End If
    End If
    modCourrier.RemplacerCorps nouveau, d("Corps")
    On Error Resume Next
    modGras.AppliquerGras nouveau
    On Error GoTo 0
    Set AssemblerLettre = nouveau
End Function
