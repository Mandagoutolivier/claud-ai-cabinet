Attribute VB_Name = "NewMacros"
Sub corrigecho()
Attribute corrigecho.VB_Description = "Macro enregistrée le 27/02/00 par personnel"
Attribute corrigecho.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.corrigecho"
'
' corrigecho Macro
' Macro enregistrée le 27/02/00 par personnel
'
    Selection.WholeStory
    Selection.Fields.Update
End Sub
Sub insdate()
Attribute insdate.VB_Description = "Macro enregistrée le 27/02/00 par personnel"
Attribute insdate.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.insdate"
'
' insdate Macro
' Macro enregistrée le 27/02/00 par personnel
'
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText text:=" "
    Selection.InsertDateTime DateTimeFormat:="jjjj j MMMM aaaa", InsertAsField _
        :=False
End Sub
Sub Macro1()
Attribute Macro1.VB_Description = "Macro enregistrée le 08/08/00 par personnel"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Macro1"
'
' Macro1 Macro
' Macro enregistrée le 08/08/00 par personnel
'
End Sub
Sub lettretype()
Attribute lettretype.VB_Description = "Macro enregistrée le 16/02/2018 par perso"
Attribute lettretype.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.lettretype"
'
' lettretype Macro
' Macro enregistrée le 16/02/2018 par perso
'
    Documents.Add Template:="C:\msoffice\Modeles\lettreom.dot", NewTemplate:= _
        False, DocumentType:=0
    Selection.MoveDown Unit:=wdLine, Count:=16
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub etiquette()
Attribute etiquette.VB_Description = "Macro enregistrée le 06/03/2018 par perso"
Attribute etiquette.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.etiquette"
'
' etiquette Macro
' Macro enregistrée le 06/03/2018 par perso
'
    Documents.Add Template:="C:\msoffice\Modeles\ETIQUETTE.dot", NewTemplate:= _
        False, DocumentType:=0
    Selection.WholeStory
    Selection.Fields.Update
End Sub
Sub echocardiogramme()
Attribute echocardiogramme.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.echocardiogramme"
'
' echocardiogramme Macro
'
'
    Documents.Add Template:= _
        "C:\Users\accueil\Documents\Modèles Office personnalisés\ECHOCARDIOGRAMME.dotm" _
        , NewTemplate:=False, DocumentType:=0
    Selection.WholeStory
    Selection.Fields.Update
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    Selection.MoveDown Unit:=wdLine, Count:=6
End Sub
Sub lettreaudocteur()
Attribute lettreaudocteur.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.lettreaudocteur"
'
' lettreaudocteur Macro
'
'
    Documents.Add Template:="C:\Mandagout\LETTRE TYPE.dotx", NewTemplate:= _
        False, DocumentType:=0
    Selection.MoveDown Unit:=wdLine, Count:=7
    Selection.TypeText text:="a"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
End Sub
Sub INVERSERDEUXMOTS()
Attribute INVERSERDEUXMOTS.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.INVERSERDEUXMOTS"
'
' INVERSERDEUXMOTS Macro
'
'
    Selection.MoveRight Unit:=wdWord, Count:=1, Extend:=wdExtend
    Selection.Cut
    Selection.MoveLeft Unit:=wdWord, Count:=1, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
End Sub
Sub COPIERNOMFICHIER()
Attribute COPIERNOMFICHIER.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.COPIERNOMFICHIER"
'
' COPIERNOMFICHIER Macro
'
'
    Selection.TypeText text:=" "
    Selection.InsertDateTime DateTimeFormat:="yyMMddhhmm", InsertAsField:=False, _
         DateLanguage:=wdFrench, CalendarType:=wdCalendarWestern, _
        InsertAsFullWidth:=False
    Selection.MoveLeft Unit:=wdWord, Count:=3, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    
End Sub
Sub ZERO()
Attribute ZERO.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.ZERO"
'
' ZERO Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "zéro "
        .Replacement.text = "0,"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub seconde()
Attribute seconde.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.seconde"
'
' seconde Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "secondes"
        .Replacement.text = "s"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub parminute()
Attribute parminute.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.parminute"
'
' parminute Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "par minute"
        .Replacement.text = "/min"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub KARDEGIC()
Attribute KARDEGIC.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.KARDEGIC"
'
' KARDEGIC Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .SmallCaps = False
        .AllCaps = True
    End With
    With Selection.Find
        .text = "Kardégic"
        .Replacement.text = "KARDEGIC"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub mmdemercure()
Attribute mmdemercure.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.mmdemercure"
'
' mmdemercure Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "mm de mercure"
        .Replacement.text = "mmHg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub signet()
Attribute signet.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.signet"
'
' signet Macro
'
'
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="nom"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub
Sub ajoutersignet()
Attribute ajoutersignet.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.ajoutersignet"
'
' ajoutersignet Macro
'
'
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="nom"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With
End Sub
Sub Majuscule()
Attribute Majuscule.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.Majuscule"
'
' Majuscule Macro
'
'
    Selection.Range.Case = wdNextCase
End Sub

Sub MiseEnFormeComplete()
    Dim para As Paragraph
    Dim doc As Document
    Set doc = ActiveDocument
    
    '---------------------------
    ' 1. Police et taille
    '---------------------------
    doc.content.Font.Name = "Times New Roman"
    doc.content.Font.Size = 10
    
    '---------------------------
    ' 2. Alinéa de 3 cm
    '---------------------------
    For Each para In doc.Paragraphs
        para.FirstLineIndent = CentimetersToPoints(3)
    Next para
    
    '---------------------------
    ' 3. Destinataire, expéditeur, date et lieu à droite
    '---------------------------
    ' Destinataire = 1er paragraphe
    doc.Paragraphs(1).Alignment = wdAlignParagraphRight
    doc.Paragraphs(1).Range.Font.Bold = True
    
    ' Expéditeur = 2e paragraphe
    doc.Paragraphs(2).Alignment = wdAlignParagraphRight
    doc.Paragraphs(2).Range.Font.Bold = True
    
    ' Date et lieu = 3e paragraphe
    doc.Paragraphs(3).Alignment = wdAlignParagraphRight
    
   
    
    MsgBox "Mise en forme complète terminée !", vbInformation
End Sub

Sub Retrait8cm()
    ' Applique un retrait gauche de 8 cm au paragraphe sélectionné
    Selection.ParagraphFormat.LeftIndent = CentimetersToPoints(8)
End Sub
Sub alinea3()
Attribute alinea3.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.alinea3"
'
' alinea3 Macro
'
'
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(3)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(3)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
    End With
End Sub
Sub alinea3racc()
Attribute alinea3racc.VB_ProcData.VB_Invoke_Func = "TemplateProject.NewMacros.alinea3racc"
'
' alinea3racc Macro
'
'
    Application.Run MacroName:="TemplateProject.NewMacros.alinea3"
End Sub
