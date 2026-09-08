Attribute VB_Name = "modCerfaPrint"
Option Explicit
' =====================================================================
' modCerfaPrint - Impression calee sur la feuille de soins pre-imprimee
' (cerfa S3110). Le document Word est construit A LA VOLEE : une zone de
' texte flottante par champ, positions en mm dans Config\cerfa_positions.txt,
' decalage global d'imprimante dans Config\cerfa_offsets.txt (dx;dy en mm).
' =====================================================================

Private Const MM_EN_POINTS As Double = 2.834645

Private Const PAGE_L_MM As Double = 210
Private Const PAGE_H_MM As Double = 297
Private Const BOITE_H_PT As Double = 18

' [CERFA] Rotation180=1 : l'image entiere est retournee de 180 degres
' (liasse qui ne peut etre inseree que dans un sens dans le bac).
Private Function Rotation180() As Boolean
    Rotation180 = (modConfig.ConfigNum("CERFA", "Rotation180", 0) = 1)
End Function

' Zone de texte a (x, y) mm (coin haut-gauche du texte), largeur w mm.
' Applique la rotation 180 si configuree : la boite est placee au point
' symetrique et son contenu retourne.
Private Function AjouterTexte(ByVal doc As Object, ByVal xmm As Double, ByVal ymm As Double, _
                              ByVal wmm As Double, ByVal taille As Double, ByVal texte As String) As Object
    Dim shp As Object, gauche As Double, haut As Double
    If Rotation180() Then
        gauche = (PAGE_L_MM - xmm - wmm) * MM_EN_POINTS
        haut = PAGE_H_MM * MM_EN_POINTS - ymm * MM_EN_POINTS - BOITE_H_PT
    Else
        gauche = xmm * MM_EN_POINTS
        haut = ymm * MM_EN_POINTS
    End If
    Set shp = doc.Shapes.AddTextbox(1, gauche, haut, wmm * MM_EN_POINTS, BOITE_H_PT)
    shp.TextFrame.TextRange.Text = texte
    shp.TextFrame.TextRange.Font.Name = "Arial"
    shp.TextFrame.TextRange.Font.Size = taille
    shp.TextFrame.MarginLeft = 0
    shp.TextFrame.MarginTop = 0
    shp.TextFrame.WordWrap = 0
    shp.Line.Visible = 0        ' msoFalse
    shp.Fill.Visible = 0
    If Rotation180() Then shp.Rotation = 180
    Set AjouterTexte = shp
End Function

' Champ a cases (dates JJMMAAAA, NIR) : un caractere par case, au pas
' 'pas' mm ; les separateurs / . espace sont retires.
Private Sub AjouterTexteCases(ByVal doc As Object, ByVal xmm As Double, ByVal ymm As Double, _
                              ByVal pas As Double, ByVal taille As Double, ByVal texte As String)
    Dim t As String, i As Long
    t = Replace(Replace(Replace(texte, "/", ""), ".", ""), " ", "")
    For i = 1 To Len(t)
        AjouterTexte doc, xmm + (i - 1) * pas, ymm, pas + 2, taille, Mid$(t, i, 1)
    Next i
End Sub

Private Sub AjouterLigne(ByVal doc As Object, ByVal x1 As Double, ByVal y1 As Double, _
                         ByVal x2 As Double, ByVal y2 As Double, ByVal epais As Double)
    Dim shp As Object
    If Rotation180() Then
        x1 = PAGE_L_MM - x1: x2 = PAGE_L_MM - x2
        y1 = PAGE_H_MM - y1: y2 = PAGE_H_MM - y2
    End If
    Set shp = doc.Shapes.AddLine(x1 * MM_EN_POINTS, y1 * MM_EN_POINTS, x2 * MM_EN_POINTS, y2 * MM_EN_POINTS)
    shp.Line.Weight = epais
    shp.Line.ForeColor.RGB = RGB(120, 120, 120)
End Sub

Private Function NouveauDocumentCerfa(ByRef word As Object) As Object
    Dim doc As Object
    Set word = CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    Set doc = word.Documents.Add()
    doc.PageSetup.PaperSize = 7           ' wdPaperA4
    doc.PageSetup.Orientation = 0         ' portrait
    doc.PageSetup.TopMargin = 0: doc.PageSetup.BottomMargin = 0
    doc.PageSetup.LeftMargin = 0: doc.PageSetup.RightMargin = 0
    Set NouveauDocumentCerfa = doc
End Function

Private Sub ImprimerOuExporter(ByVal word As Object, ByVal doc As Object, ByVal versPdf As String)
    If Len(versPdf) > 0 Then
        doc.ExportAsFixedFormat versPdf, 17
    Else
        Dim imprimante As String
        imprimante = modConfig.Config("CERFA", "Imprimante", "")
        If Len(imprimante) > 0 Then word.ActivePrinter = imprimante
        doc.PrintOut Background:=False
    End If
End Sub

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
    Dim word As Object, doc As Object
    Dim positions As Collection, p As Variant, dx As Double, dy As Double

    Set positions = LirePositions()
    LireOffsets dx, dy
    On Error GoTo Nettoyage
    Set doc = NouveauDocumentCerfa(word)
    For Each p In positions
        If valeurs.Exists(CStr(p(0))) Then
            If Len(CStr(valeurs(p(0)))) > 0 Then
                If CDbl(p(5)) > 0 Then
                    AjouterTexteCases doc, CDbl(p(1)) + dx, CDbl(p(2)) + dy, CDbl(p(5)), CDbl(p(4)), CStr(valeurs(p(0)))
                Else
                    AjouterTexte doc, CDbl(p(1)) + dx, CDbl(p(2)) + dy, CDbl(p(3)), CDbl(p(4)), CStr(valeurs(p(0)))
                End If
            End If
        End If
    Next p
    ImprimerOuExporter word, doc, versPdf
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

' Impression de croix de reperes puis saisie des ecarts mesures.
' Les croix sont imprimees AVEC le decalage deja enregistre (comme les
' feuilles de soins) : on peut donc relancer le calage jusqu'a ce que les
' croix tombent juste ; chaque mesure s'AJOUTE au decalage en vigueur.
Public Sub CalageCerfa()
    Dim reponse As String, dx As Double, dy As Double, ddx As Double, ddy As Double
    LireOffsets dx, dy
    If dx <> 0 Or dy <> 0 Then
        Select Case MsgBox("Decalage actuellement enregistre : dx=" & dx & " mm, dy=" & dy & " mm." & vbCrLf & vbCrLf & _
                           "Oui = repartir de ZERO (croix sans decalage)" & vbCrLf & _
                           "Non = garder ce decalage et l'affiner" & vbCrLf & _
                           "Annuler = quitter", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Calage CERFA")
            Case vbYes: dx = 0: dy = 0
            Case vbCancel: Exit Sub
        End Select
    End If
    If MsgBox("Placez une feuille de soins SACRIFIEE dans l'imprimante." & vbCrLf & _
              "Des croix de repere vont s'imprimer aux 4 coins (a 20 mm des bords)," & vbCrLf & _
              "avec le decalage en vigueur (dx=" & dx & " mm, dy=" & dy & " mm)." & vbCrLf & vbCrLf & _
              "Imprimer maintenant ?", vbOKCancel + vbInformation, "Calage CERFA") <> vbOK Then Exit Sub

    Dim word As Object, doc As Object, coords As Variant, c As Variant
    On Error GoTo Nettoyage
    Set doc = NouveauDocumentCerfa(word)
    coords = Array(Array(20, 20), Array(190, 20), Array(20, 277), Array(190, 277))
    For Each c In coords
        AjouterLigne doc, CDbl(c(0)) - 5 + dx, CDbl(c(1)) + dy, CDbl(c(0)) + 5 + dx, CDbl(c(1)) + dy, 0.5
        AjouterLigne doc, CDbl(c(0)) + dx, CDbl(c(1)) - 5 + dy, CDbl(c(0)) + dx, CDbl(c(1)) + 5 + dy, 0.5
    Next c
    AjouterTexte doc, 60 + dx, 8 + dy, 90, 8, "HAUT de la feuille - calage CERFA " & Format$(Now, "dd/mm hh:nn")
    ImprimerOuExporter word, doc, ""
    doc.Close 0
    word.Quit
    On Error GoTo 0

    reponse = InputBox("Mesurez sur la feuille imprimee, croix HAUT-GAUCHE :" & vbCrLf & _
        "ecart HORIZONTAL en mm par rapport a 20 mm du bord gauche" & vbCrLf & _
        "(positif si la croix est trop a DROITE, negatif si trop a gauche ; 0 si juste)", _
        "Calage CERFA - decalage X", "0")
    If Len(reponse) = 0 Then Exit Sub
    ddx = -Val(Replace(reponse, ",", "."))
    reponse = InputBox("ecart VERTICAL en mm par rapport a 20 mm du bord haut" & vbCrLf & _
        "(positif si la croix est trop BASSE, negatif si trop haute ; 0 si juste)", _
        "Calage CERFA - decalage Y", "0")
    If Len(reponse) = 0 Then Exit Sub
    ddy = -Val(Replace(reponse, ",", "."))
    dx = dx + ddx: dy = dy + ddy
    modFichiers.EcrireTexteUTF8 modConfig.Chemin("Config") & "\cerfa_offsets.txt", _
        Replace(CStr(dx), ",", ".") & ";" & Replace(CStr(dy), ",", ".")
    MsgBox "Calage enregistre : dx=" & dx & " mm, dy=" & dy & " mm" & _
           IIf(ddx <> 0 Or ddy <> 0, " (correction de " & ddx & " / " & ddy & " mm)", "") & "." & vbCrLf & _
           "Relancez le calage pour verifier : les croix doivent tomber a 20 mm des bords.", vbInformation, "Calage CERFA"
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


' Grille de calage : lignes tous les 10 mm (fines tous les 5 mm) graduees
' en mm depuis le coin HAUT-GAUCHE, et le NOM de chaque champ imprime a sa
' position actuelle. Sur une feuille de soins sacrifiee, on lit directement
' les coordonnees x;y de chaque case et on corrige cerfa_positions.txt.
Public Sub CalageCerfaGrille()
    Dim word As Object, doc As Object, positions As Collection, p As Variant
    Dim dx As Double, dy As Double, i As Long
    If MsgBox("Placez une feuille de soins SACRIFIEE dans l'imprimante." & vbCrLf & _
              "Une grille graduee en mm (origine = coin haut-gauche) et le nom de chaque" & vbCrLf & _
              "champ a sa position actuelle vont s'imprimer, avec le decalage en vigueur" & vbCrLf & _
              IIf(Rotation180(), "et la rotation de 180 degres ([CERFA] Rotation180=1).", "([CERFA] Rotation180=0).") & vbCrLf & vbCrLf & _
              "Imprimer maintenant ?", vbOKCancel + vbInformation, "Grille de calage CERFA") <> vbOK Then Exit Sub
    LireOffsets dx, dy
    On Error GoTo Nettoyage
    Set doc = NouveauDocumentCerfa(word)
    For i = 0 To 210 Step 5
        AjouterLigne doc, i + dx, dy, i + dx, 297 + dy, IIf(i Mod 10 = 0, 0.5, 0.25)
        If i Mod 10 = 0 And i > 0 Then AjouterTexte doc, i + 0.5 + dx, 1 + dy, 9, 6, CStr(i)
    Next i
    For i = 0 To 297 Step 5
        AjouterLigne doc, dx, i + dy, 210 + dx, i + dy, IIf(i Mod 10 = 0, 0.5, 0.25)
        If i Mod 10 = 0 And i > 0 Then AjouterTexte doc, 1 + dx, i + 0.5 + dy, 9, 6, CStr(i)
    Next i
    Set positions = LirePositions()
    For Each p In positions
        AjouterTexte doc, CDbl(p(1)) + dx, CDbl(p(2)) + dy, CDbl(p(3)), CDbl(p(4)), CStr(p(0))
    Next p
    AjouterTexte doc, 40 + dx, 290 + dy, 130, 7, "Grille CERFA " & Format$(Now, "dd/mm/yyyy hh:nn") & _
        " - dx=" & dx & " dy=" & dy & " - Rotation180=" & IIf(Rotation180(), "1", "0")
    ImprimerOuExporter word, doc, ""
    doc.Close 0
    word.Quit
    MsgBox "Grille imprimee. Pour chaque case de la liasse, lisez x (mm depuis la gauche) et y" & vbCrLf & _
           "(mm depuis le haut) du coin haut-gauche de la case, puis reportez-les dans" & vbCrLf & _
           modConfig.Chemin("Config") & "\cerfa_positions.txt (CHAMP;x;y;largeur;police)." & vbCrLf & vbCrLf & _
           "Si tout est a l'envers : mettez [CERFA] Rotation180=1 dans config.ini." & vbCrLf & _
           "Si tout est decale d'un meme ecart : bouton 'Calage imprimante CERFA'.", vbInformation, "Grille de calage CERFA"
    Exit Sub
Nettoyage:
    Dim descErr As String
    descErr = Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    word.Quit
    On Error GoTo 0
    MsgBox "Erreur d'impression : " & descErr, vbCritical, "Grille de calage CERFA"
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
                Dim pas As Double
                pas = 0
                If UBound(parties) >= 5 Then pas = Val(Replace(parties(5), ",", "."))
                col.Add Array(Trim$(parties(0)), Val(Replace(parties(1), ",", ".")), _
                              Val(Replace(parties(2), ",", ".")), Val(Replace(parties(3), ",", ".")), _
                              Val(Replace(parties(4), ",", ".")), pas)
            End If
        End If
    Next i
    Set LirePositions = col
End Function

' Decalage global : Config\cerfa_offsets.txt (ecrit par le calage) ;
' a defaut, [CERFA] OffsetX_mm / OffsetY_mm de config.ini.
Private Sub LireOffsets(ByRef dx As Double, ByRef dy As Double)
    Dim chemin As String, contenu As String, parties() As String
    dx = modConfig.ConfigNum("CERFA", "OffsetX_mm", 0)
    dy = modConfig.ConfigNum("CERFA", "OffsetY_mm", 0)
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
