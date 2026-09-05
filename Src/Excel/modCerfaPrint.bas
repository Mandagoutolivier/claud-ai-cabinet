Attribute VB_Name = "modCerfaPrint"
Option Explicit
' =====================================================================
' modCerfaPrint - Impression calee sur la feuille de soins pre-imprimee
' (cerfa S3110). Le document Word est construit A LA VOLEE : une zone de
' texte flottante par champ, positions en mm dans Config\cerfa_positions.txt,
' decalage global d'imprimante dans Config\cerfa_offsets.txt (dx;dy en mm).
' =====================================================================

Private Const MM_EN_POINTS As Double = 2.834645

' Remplit et imprime la feuille de soins papier.
' infos : dictionnaire de la seance. Le PATIENT qui recoit les soins et
'   l'ASSURE peuvent etre deux personnes differentes : le patient vient de
'   Nom/Prenom/DDN/NIR, l'assure de AssureNom/AssurePrenom/AssureDDN/AssureNIR
'   et, a defaut, le patient est son propre assure.
'   MedTraitantNom : a renseigner quand le medecin intervient comme
'   correspondant d'un patient adresse (rubrique parcours de soins).
' actes : Collection de dictionnaires (CodeActe/Code, Montant/Tarif, DateActe)
'   La DATE portee sur chaque ligne est celle de la REALISATION de l'acte,
'   jamais la date du jour d'impression.
' versPdf : si renseigne, exporte en PDF au lieu d'imprimer. Ce PDF ne
'   reconstitue pas le fond du Cerfa : il sert au controle du calage.
' Renvoie True quand le document a ete envoye a l'imprimante. Cela ne
' prouve pas que la feuille est physiquement sortie : la confirmation
' de la sortie papier est demandee a la secretaire par l'appelant.
Public Function ImprimerFeuille(ByVal infos As Object, ByVal actes As Collection, _
                                Optional ByVal versPdf As String = "") As Boolean
    Dim valeurs As Object, a As Object, i As Long, total As Double
    Dim code As String, montant As String, dateLigne As String
    Dim maxLignes As Long

    maxLignes = CLng(modConfig.ConfigNum("CERFA", "LignesMax", 4))
    If maxLignes < 1 Then maxLignes = 4
    If actes.Count > maxLignes Then
        ' refus explicite : aucune perte silencieuse d'acte ni de total
        Err.Raise vbObjectError + 701, "modCerfaPrint", _
            actes.Count & " actes pour une feuille de soins qui n'en porte que " & maxLignes & "." & vbCrLf & _
            "Etablissez deux feuilles (traitez la seance en deux fois) ou corrigez la selection : " & _
            "le logiciel n'imprime pas une feuille incomplete."
    End If

    Set valeurs = CreateObject("Scripting.Dictionary")
    valeurs.CompareMode = 1

    ' --- patient qui recoit les soins ---
    valeurs("PATIENT_NOM") = Trim$(ValeurOuVide(infos, "Nom") & " " & ValeurOuVide(infos, "Prenom"))
    valeurs("PATIENT_DDN") = ValeurOuVide(infos, "DDN")
    valeurs("PATIENT_NIR") = ValeurOuVide(infos, "NIR")

    ' --- assure : distinct s'il est renseigne, sinon le patient lui-meme ---
    If Len(ValeurOuVide(infos, "AssureNom")) > 0 Then
        valeurs("ASSURE_NOM") = Trim$(ValeurOuVide(infos, "AssureNom") & " " & ValeurOuVide(infos, "AssurePrenom"))
        valeurs("ASSURE_DDN") = ValeurOuVide(infos, "AssureDDN")
        valeurs("ASSURE_NIR") = ValeurOuVide(infos, "AssureNIR")
    Else
        valeurs("ASSURE_NOM") = valeurs("PATIENT_NOM")
        valeurs("ASSURE_DDN") = valeurs("PATIENT_DDN")
        valeurs("ASSURE_NIR") = valeurs("PATIENT_NIR")
    End If

    ' --- parcours de soins : medecin traitant du patient adresse ---
    valeurs("MEDECIN_TRAITANT") = ValeurOuVide(infos, "MedTraitantNom")

    i = 0
    For Each a In actes
        i = i + 1
        If a.Exists("CodeActe") Then code = a("CodeActe") Else code = ValeurOuVide(a, "Code")
        If a.Exists("Montant") Then montant = a("Montant") Else montant = ValeurOuVide(a, "Tarif")
        dateLigne = ValeurOuVide(a, "DateActe")
        If Len(dateLigne) <> 10 Then dateLigne = ValeurOuVide(infos, "DateActe")
        If Len(dateLigne) <> 10 Then
            Err.Raise vbObjectError + 702, "modCerfaPrint", _
                "Date de realisation manquante pour l'acte " & code & " : " & _
                "la feuille de soins ne peut pas porter la date du jour a sa place."
        End If
        valeurs("DATE" & i) = dateLigne
        valeurs("CODE" & i) = code
        valeurs("MONTANT" & i) = Format$(Val(Replace(montant, ",", ".")), "0.00")
        total = total + Val(Replace(montant, ",", "."))
    Next a
    valeurs("TOTAL") = Format$(total, "0.00")

    ImprimerDocumentCale valeurs, versPdf
    ImprimerFeuille = True
End Function

' Construit et imprime (ou exporte) le document cale
Private Sub ImprimerDocumentCale(ByVal valeurs As Object, ByVal versPdf As String)
    Dim word As Object, doc As Object, shp As Object
    Dim positions As Collection, p As Variant, dx As Double, dy As Double

    Set positions = LirePositions()
    LireOffsets dx, dy

    Set word = CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    On Error GoTo Nettoyage
    Set doc = word.Documents.Add()
    doc.PageSetup.TopMargin = 0
    doc.PageSetup.BottomMargin = 0
    doc.PageSetup.LeftMargin = 0
    doc.PageSetup.RightMargin = 0

    For Each p In positions
        If valeurs.Exists(CStr(p(0))) Then
            If Len(CStr(valeurs(p(0)))) > 0 Then
                Set shp = doc.Shapes.AddTextbox(1, _
                    (CDbl(p(1)) + dx) * MM_EN_POINTS, (CDbl(p(2)) + dy) * MM_EN_POINTS, _
                    CDbl(p(3)) * MM_EN_POINTS, 18)
                shp.TextFrame.TextRange.Text = CStr(valeurs(p(0)))
                shp.TextFrame.TextRange.Font.Name = "Arial"
                shp.TextFrame.TextRange.Font.Size = CDbl(p(4))
                shp.TextFrame.MarginLeft = 0
                shp.TextFrame.MarginTop = 0
                shp.Line.Visible = 0        ' msoFalse
                shp.Fill.Visible = 0
            End If
        End If
    Next p

    If Len(versPdf) > 0 Then
        doc.ExportAsFixedFormat versPdf, 17
    Else
        Dim imprimante As String
        imprimante = modConfig.Config("CERFA", "Imprimante", "")
        If Len(imprimante) > 0 Then word.ActivePrinter = imprimante
        doc.PrintOut Background:=False
    End If
    doc.Close 0
    word.Quit
    Exit Sub
Nettoyage:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit
    On Error GoTo 0
    Err.Raise numErr, "modCerfaPrint", descErr
End Sub

' Impression de croix de reperes puis saisie des ecarts mesures
Public Sub CalageCerfa()
    Dim valeurs As Object, reponse As String, dx As Double, dy As Double
    If MsgBox("Placez une feuille de soins SACRIFIEE dans l'imprimante." & vbCrLf & _
              "Des croix de repere vont s'imprimer aux 4 coins (a 20 mm des bords)." & vbCrLf & vbCrLf & _
              "Imprimer maintenant ?", vbOKCancel + vbInformation, "Calage CERFA") <> vbOK Then Exit Sub

    Dim word As Object, doc As Object, shp As Object, coords As Variant, c As Variant
    Set word = CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    On Error GoTo Nettoyage
    Set doc = word.Documents.Add()
    doc.PageSetup.TopMargin = 0: doc.PageSetup.BottomMargin = 0
    doc.PageSetup.LeftMargin = 0: doc.PageSetup.RightMargin = 0
    coords = Array(Array(20, 20), Array(190, 20), Array(20, 277), Array(190, 277))
    For Each c In coords
        Set shp = doc.Shapes.AddTextbox(1, (CDbl(c(0)) - 2) * MM_EN_POINTS, _
                                        (CDbl(c(1)) - 4) * MM_EN_POINTS, 12 * MM_EN_POINTS, 16)
        shp.TextFrame.TextRange.Text = "+"
        shp.TextFrame.TextRange.Font.Size = 14
        shp.TextFrame.MarginLeft = 0: shp.TextFrame.MarginTop = 0
        shp.Line.Visible = 0: shp.Fill.Visible = 0
    Next c
    Dim imprimante As String
    imprimante = modConfig.Config("CERFA", "Imprimante", "")
    If Len(imprimante) > 0 Then word.ActivePrinter = imprimante
    doc.PrintOut Background:=False
    doc.Close 0
    word.Quit
    On Error GoTo 0

    reponse = InputBox("Mesurez sur la feuille imprimee :" & vbCrLf & _
        "- ecart HORIZONTAL en mm entre la croix haut-gauche et 20 mm du bord gauche" & vbCrLf & _
        "  (positif si la croix est trop a droite)", "Calage CERFA - decalage X", "0")
    If Len(reponse) = 0 Then Exit Sub
    dx = -Val(Replace(reponse, ",", "."))
    reponse = InputBox("- ecart VERTICAL en mm entre la croix haut-gauche et 20 mm du bord haut" & vbCrLf & _
        "  (positif si la croix est trop basse)", "Calage CERFA - decalage Y", "0")
    If Len(reponse) = 0 Then Exit Sub
    dy = -Val(Replace(reponse, ",", "."))
    modFichiers.EcrireTexteUTF8 modConfig.Chemin("Config") & "\cerfa_offsets.txt", _
        Replace(CStr(dx), ",", ".") & ";" & Replace(CStr(dy), ",", ".")
    MsgBox "Calage enregistre (dx=" & dx & " mm, dy=" & dy & " mm)." & vbCrLf & _
           "Refaites une impression d'essai pour verifier.", vbInformation, "Calage CERFA"
    Exit Sub
Nettoyage:
    Dim descErr As String
    descErr = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit
    On Error GoTo 0
    MsgBox "Erreur d'impression : " & descErr, vbCritical, "Calage CERFA"
End Sub

Private Function LirePositions() As Collection
    Dim chemin As String, contenu As String, lignes() As String, i As Long
    Dim parties() As String, col As Collection
    Set col = New Collection
    chemin = modConfig.Chemin("Config") & "\cerfa_positions.txt"
    If Not modFichiers.FichierExiste(chemin) Then
        Err.Raise vbObjectError + 700, "modCerfaPrint", "Fichier introuvable : " & chemin
    End If
    contenu = Replace(modFichiers.LireTexteUTF8(chemin), vbCrLf, vbLf)
    lignes = Split(contenu, vbLf)
    For i = LBound(lignes) To UBound(lignes)
        If Len(Trim$(lignes(i))) > 0 And Left$(Trim$(lignes(i)), 1) <> "#" Then
            parties = Split(lignes(i), ";")
            If UBound(parties) >= 4 Then
                col.Add Array(Trim$(parties(0)), Val(Replace(parties(1), ",", ".")), _
                              Val(Replace(parties(2), ",", ".")), Val(Replace(parties(3), ",", ".")), _
                              Val(Replace(parties(4), ",", ".")))
            End If
        End If
    Next i
    Set LirePositions = col
End Function

Private Sub LireOffsets(ByRef dx As Double, ByRef dy As Double)
    Dim chemin As String, contenu As String, parties() As String
    dx = 0: dy = 0
    chemin = modConfig.Chemin("Config") & "\cerfa_offsets.txt"
    If Not modFichiers.FichierExiste(chemin) Then Exit Sub
    contenu = Trim$(Replace(Replace(modFichiers.LireTexteUTF8(chemin), vbCr, ""), vbLf, ""))
    parties = Split(contenu, ";")
    If UBound(parties) >= 1 Then
        dx = Val(Replace(parties(0), ",", "."))
        dy = Val(Replace(parties(1), ",", "."))
    End If
End Sub

Private Function ValeurOuVide(ByVal dict As Object, ByVal cle As String) As String
    If dict.Exists(cle) Then ValeurOuVide = dict(cle) Else ValeurOuVide = ""
End Function
