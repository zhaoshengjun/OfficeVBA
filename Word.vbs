'******************************
' Special part
'******************************
Const fontEN = "Lato"
Const normalFontSize = 12
Const lineSpace = 16
Const fontHeaderEN = "Lato"
'Const fontEN = "Eureka Sans"
'Const normalFontSize = 16
'Const lineSpace = 20
'******************************
' Common part
'******************************
Const fontCN = "Arial"
Const headerFontSize = 28
Private Sub replaceEnterKey()
'?????
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub changeFonts()
'????????
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    Selection.WholeStory
    Selection.Font.Name = fontCN
    Selection.Font.Name = fontEN
    Selection.Font.Size = normalFontSize
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = lineSpace
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0.35)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 2
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub
Private Sub replaceBlankLine()
' ???????
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub clearFormatOfPictures()
    On Error Resume Next
    For i = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(i).Select
        Selection.ClearFormatting
        Call formatPicture
    Next i
End Sub
Sub removeLinks()
    With ActiveDocument
        Dim myLink As Hyperlink
        Dim myBookmark As Bookmark
        Dim myField As Field
        For Each myLink In .Hyperlinks
            myLink.Delete '???????
        Next myLink
        For Each myBookmark In .Bookmarks
            myBookmark.Delete '??"??"??"??"(???????)
        Next myBookmark
        For Each myField In .Fields
            myField.Unlink  '????????
        Next myField
    End With
End Sub
Private Sub changeFontEN()
    ActiveWindow.ActivePane.VerticalPercentScrolled = 0
    Selection.WholeStory
    Selection.Font.Name = fontEN
    Selection.Font.Size = normalFontSize
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = lineSpace
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
'        .MirrorIndents = False
'        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub
Sub formatAll()
    Call replaceEnterKey
    Call changeFonts
    Call replaceBlankLine
    Call replaceBlankLine
    Call replaceBlankLine
    Call clearFormatOfPictures
    Call formatHeader
    Call formatList
'    Call removeLinks
End Sub
Sub formatAllEN()
    Call replaceEnterKey
    Call changeFontEN
    Call replaceBlankLine
    Call replaceBlankLine
    Call replaceBlankLine
    Call clearFormatOfPictures
    Call formatHeader
    Call formatList
    Call formatLinks
End Sub

Sub formatName()
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ".bak"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = "_"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = "-"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub formatHeader()
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.ClearFormatting
'    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.Font.Name = fontCN
    Selection.Font.Name = fontHeaderEN
    Selection.Font.Size = headerFontSize
    Selection.Font.Bold = wdToggle
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
    End With
End Sub
Sub saveAsPDF()
    aFileName = ActiveDocument.FullName
    If aFileName <> "" Then
    aFileName = Mid(aFileName, 1, Len(aFileName) - 4) & ".pdf"
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        aFileName, ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:= _
        wdExportAllDocument, From:=1, To:=1, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:= _
        wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:= _
        True, UseISO19005_1:=False
    End If
End Sub
'Sub AutoClose()
'    If ActiveDocument.FullName <> "" Then
'        Call saveAsPDF
'    End If
'End Sub
Sub formatNote()
'    Call createBlankLines
    For Each aPara In ActiveDocument.Paragraphs
        If Len(aPara.Range.Text) > 1 Then
            aPara.Range.Select
            Call formatAsQuote
        End If
    Next aPara
End Sub

Private Sub createBlankLines()
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Private Sub formatAsQuote()
    With Selection.ParagraphFormat
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth450pt
            .Color = -721354753
        End With
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth450pt
        .DefaultBorderColor = -721354753
    End With
End Sub
Private Sub formatPicture()

    Selection.InlineShapes(1).Fill.Visible = msoFalse
    Selection.InlineShapes(1).Fill.Solid
    Selection.InlineShapes(1).Fill.Transparency = 0#
    Selection.InlineShapes(1).Line.Weight = 0.75
    Selection.InlineShapes(1).Line.Transparency = 0#
    Selection.InlineShapes(1).Line.Visible = msoFalse
    Selection.InlineShapes(1).LockAspectRatio = msoTrue
    Selection.InlineShapes(1).Height = 269.3
    Selection.InlineShapes(1).Width = 414.7
    Selection.InlineShapes(1).PictureFormat.Brightness = 0.5
    Selection.InlineShapes(1).PictureFormat.Contrast = 0.5
    Selection.InlineShapes(1).PictureFormat.ColorType = msoPictureAutomatic
    Selection.InlineShapes(1).PictureFormat.CropLeft = 0#
    Selection.InlineShapes(1).PictureFormat.CropRight = 0#
    Selection.InlineShapes(1).PictureFormat.CropTop = 0#
    Selection.InlineShapes(1).PictureFormat.CropBottom = 0#
End Sub
Private Sub formatList()
    For Each aList In ActiveDocument.Lists
        aList.Range.Select
        Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(-0.5)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(-0.25)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
        End With
        With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
            .NumberFormat = ChrW(61548)
            .TrailingCharacter = wdTrailingTab
            .NumberStyle = wdListNumberStyleBullet
            .NumberPosition = CentimetersToPoints(0)
            .Alignment = wdListLevelAlignLeft
            .TextPosition = CentimetersToPoints(0.74)
            .TabPosition = wdUndefined
            .ResetOnHigher = 0
            .StartAt = 1
            With .Font
                .Bold = wdUndefined
                .Italic = wdUndefined
                .StrikeThrough = wdUndefined
                .Subscript = wdUndefined
                .Superscript = wdUndefined
                .Shadow = wdUndefined
                .Outline = wdUndefined
                .Emboss = wdUndefined
                .Engrave = wdUndefined
                .AllCaps = wdUndefined
                .Hidden = wdUndefined
                .Underline = wdUndefined
                .Color = wdUndefined
                .Size = wdUndefined
                .Animation = wdUndefined
                .DoubleStrikeThrough = wdUndefined
                .Name = "Wingdings"
            End With
            .LinkedStyle = ""
        End With
        ListGalleries(wdBulletGallery).ListTemplates(1).Name = ""
'        Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
'            ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
'            False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
'            wdWord10ListBehavior
        Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=ListGalleries( _
        wdBulletGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:= _
        wdListApplyToWholeList, DefaultListBehavior:=wdWord10ListBehavior
    Next aList
End Sub


Sub formatFIReport()
'***********************************************************
' This macro will format the fi report copied from PDF file
' into correct style with tab separated columns.
'***********************************************************
    Selection.WholeStory
    Selection.Font.Name = "Courier New"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "("
        .Replacement.Text = "-"
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = ")"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = " "
        .Replacement.Text = vbTab
        .Forward = True
        .Wrap = wdFindAsk
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     With Selection.Find
        .Text = ","
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Call textFormat
End Sub

Private Sub textFormat()
    Set myrange = ActiveDocument.Range(Start:=0, End:=Selection.End)
    Index = 0
    For Each aWord In myrange.Words
        If Not IsNumeric(aWord.Text) Then
            If Index > 0 And aWord.Text <> "-" Then
                If myrange.Words(Index).Text = vbTab Then
                    myrange.Words(Index).Text = " "
                    Index = Index - 1
                End If
            End If
        End If
        Index = Index + 1
    Next aWord
End Sub
Sub convert2HTML()
    Selection.WholeStory
    Set myrange = ActiveDocument.Range(Start:=0, End:=Selection.End)
'    Set myRange = Selection.Range
    For Each aLink In myrange.Hyperlinks
        aLinkAddress = aLink.Address
        aLinkAddress = " <a href=""" & aLinkAddress & """>"
'        aLinkRange = ActiveDocument.Range(aLink.Range.Start, aLink.Range.End)
        aLink.Range.InsertBefore aLinkAddress
        aLink.Range.InsertAfter "</a>"
        aLink.Range.Select
        Selection.ClearFormatting
    Next aLink

    For Each aPara In myrange.Paragraphs

        If Len(aPara.Range.Text) > 2 Then
            Set aRange = ActiveDocument.Range(aPara.Range.Start, aPara.Range.End - 1)
            If aRange.ListFormat.ListValue <> 0 Then
                aRange.InsertBefore "<ul><li>"
                aRange.InsertAfter "</ul></li>"
            Else
                Select Case aPara.Style.NameLocal
                    Case "Heading 1"
                        aRange.InsertBefore "<h1>"
                        aRange.InsertAfter "</h1>"
                    Case "Heading 2"
                        aRange.InsertBefore "<h2>"
                        aRange.InsertAfter "</h2>"
                    Case "Heading 3"
                        aRange.InsertBefore "<h3>"
                        aRange.InsertAfter "</h3>"
                    Case "Heading 4"
                        aRange.InsertBefore "<h4>"
                        aRange.InsertAfter "</h4>"
                    Case Else
                        aRange.InsertBefore "<p>"
                        aRange.InsertAfter "</p>"
                End Select
            End If
        End If
    Next aPara
    Selection.WholeStory
    Selection.ClearFormatting
End Sub

Sub formatLinks()
    With ActiveDocument
        Dim myLink As Hyperlink
        Dim myBookmark As Bookmark
        Dim myField As Field
        For Each myLink In .Hyperlinks
            myLink.Range.Select
            Call formatLinkFont
        Next myLink
        For Each myBookmark In .Bookmarks
            myBookmark.Range.Select
            Call formatLinkFont
        Next myBookmark
'        For Each myField In .Fields
'            myField.Select
'            Selection.Font.Name = "Courier New"
'        Next myField
    End With
End Sub

Sub formatTables()
    With ActiveDocument
        Dim myTable As Table
        For Each myTable In .Tables
            myTable.Range.Select
            Selection.Tables(1).PreferredWidth = CentimetersToPoints(16)
        Next myTable
    End With
End Sub
Sub formatLinkFont()
    With Selection.Font
        .Name = "Lato"
        .Size = 12
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineThick
        .UnderlineColor = wdColorOrange
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorBlue
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
End Sub
Sub formatFont()

    For Each wd In ActiveDocument.Words
        wd.Select
        If Selection.Font.NameBi <> "Courier New" And _
           Selection.Font.NameBi <> "Consolas" Then
            With Selection.Font
                .Name = "Lato"
                If .Size < 14 Then
                    .Size = 13
                End If
            End With
        Else
            With Selection.Font
                .Name = "Courier New"
                .Size = 10
            End With
        End If
    Next wd
End Sub
Sub minFormat()
    Call formatFont
    Call formatTables
End Sub
Sub formatCodeFont()
    With Selection.Font
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorGray05
        End With
        With .Borders(1)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorGray25
        End With
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorGray25
    End With
End Sub
