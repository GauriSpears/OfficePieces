Sub РазбитьНаСтраницы()
    Dim docMultiple As Document
    Dim docSingle As Document
    Dim rngPage As Range
    Dim iCurrentPage As Integer
    Dim iPageCount As Integer
    Dim strNewFileName As String
    Application.ScreenUpdating = False 'Makes the code run faster and reduces screen _
    flicker a bit.
    Set docMultiple = ActiveDocument 'Work on the active document _
    (the one currently containing the Selection)
    Set rngPage = docMultiple.Range 'instantiate the range object
    iCurrentPage = 1
    'По сколько страниц разбивать
    iStep = 1
    'get the document's page count
    iPageCount = docMultiple.Content.ComputeStatistics(wdStatisticPages)
    Do Until iCurrentPage > iPageCount
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=iCurrentPage
        rngPage.Start = Selection.Start
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=iCurrentPage + iStep - 1
        rngPage.End = Selection.Bookmarks("\Page").Range.End
        rngPage.Copy 'copy the page into the Windows clipboard
        Set docSingle = Documents.Add 'create a new document
        docSingle.Range.PasteAndFormat (wdFormatOriginalFormatting) 'paste the clipboard contents to the new document
         ' Если на конце разрыв раздела (любой вид, у всех код 12).
        Dim s1 As Section
        Dim s2 As Section
        If Asc(docSingle.Characters.Last.Text) = 12 Then
            SecNum = docSingle.Sections.Count
            If SecNum >= 2 Then
                Set s1 = docSingle.Sections(SecNum - 1)
                Set s2 = docSingle.Sections(SecNum)
                PrepareToDeleteSectionBreak s1, s2
                Set s1 = Nothing
                Set s2 = Nothing
            End If
            docSingle.Characters.Last.Delete
        ' Если на конце разрыв страницы.
        ElseIf Asc(docSingle.Characters(docSingle.Characters.Count - 1).Text) = 12 Then
            SecNum = docSingle.Sections.Count
            If SecNum >= 2 Then
                Set s1 = docSingle.Sections(SecNum - 1)
                Set s2 = docSingle.Sections(SecNum)
                PrepareToDeleteSectionBreak s1, s2
                Set s1 = Nothing
                Set s2 = Nothing
            End If
            docSingle.Characters(docSingle.Characters.Count - 1).Delete
        End If
        'build a new sequentially-numbered file name based on the original multi-paged file name and path
        'strNewFileName = Replace(docMultiple.FullName, ".doc", "_" & Right$("000" & ((iCurrentPage-1)/iStep+1), 4) & ".doc")
        strNewFileName = docMultiple.Path & "\" & Replace(Trim(Replace(Replace(docSingle.Tables(3).Cell(2, 1).Range.Text, Chr(7), ""), Chr(13), "")) & " " & Trim(Replace(Replace(docSingle.Tables(2).Cell(2, 2).Range.Text, Chr(7), ""), Chr(13), "")) & ".docx", "/", "-")
        If strNewFileName <> " .docx" And strNewFileName <> docMultiple.Path & "\ .docx" Then
            docSingle.SaveAs strNewFileName 'save the new single-paged document
        End If
        iCurrentPage = iCurrentPage + iStep 'move to the next page
        docSingle.Close False 'close the new document
        'rngPage.Collapse wdCollapseEnd 'go to the next page
    Loop 'go to the top of the do loop
    Application.ScreenUpdating = True 'restore the screen updating
    'Destroy the objects.
    Set docMultiple = Nothing
    Set docSingle = Nothing
    Set rngPage = Nothing
End Sub

' PrepareToDeleteSectionBreak()
' Sets the following section's
' (and section's child objects') properties
' equal to the current section's,
' so that the break can be deleted
' without losing the current section's
' formatting.

Private Sub PrepareToDeleteSectionBreak(s1 As Section, s2 As Section)
    DuplicatePageSetupProperties s1, s2
    DuplicateColumnProperties s1, s2
    DuplicateBorderProperties s1, s2
    DuplicateHeadersAndFooters s1, s2
    DuplicatePageNumbers s1, s2
End Sub

Private Sub DuplicatePageSetupProperties(s1 As Section, s2 As Section)
    With s2.PageSetup
        ' first set up the size properties (some other properties depend on these)
        .Orientation = s1.PageSetup.Orientation
        .PageHeight = s1.PageSetup.PageHeight
        .PageWidth = s1.PageSetup.PageWidth
        
        .TopMargin = s1.PageSetup.TopMargin
        .BottomMargin = s1.PageSetup.BottomMargin
        .LeftMargin = s1.PageSetup.LeftMargin
        .RightMargin = s1.PageSetup.RightMargin
        .FooterDistance = s1.PageSetup.FooterDistance
        .HeaderDistance = s1.PageSetup.HeaderDistance
        .MirrorMargins = s1.PageSetup.MirrorMargins
        
        .VerticalAlignment = s1.PageSetup.VerticalAlignment
        
        .Gutter = s1.PageSetup.Gutter
        .GutterPos = s1.PageSetup.GutterPos
        .GutterStyle = s1.PageSetup.GutterStyle
        
        .FirstPageTray = s1.PageSetup.FirstPageTray
        .OtherPagesTray = s1.PageSetup.OtherPagesTray
        .LineNumbering = s1.PageSetup.LineNumbering
        .SectionDirection = s1.PageSetup.SectionDirection
        .SuppressEndnotes = s1.PageSetup.SuppressEndnotes
        .TwoPagesOnOne = s1.PageSetup.TwoPagesOnOne
        
        .DifferentFirstPageHeaderFooter = s1.PageSetup.DifferentFirstPageHeaderFooter
        .OddAndEvenPagesHeaderFooter = s1.PageSetup.OddAndEvenPagesHeaderFooter
        
        .SectionStart = s1.PageSetup.SectionStart
    End With
End Sub

Private Sub DuplicateColumnProperties(s1 As Section, s2 As Section)
    Dim i As Long
    With s2.PageSetup.TextColumns
        .SetCount s1.PageSetup.TextColumns.Count
        .EvenlySpaced = s1.PageSetup.TextColumns.EvenlySpaced
        .FlowDirection = s1.PageSetup.TextColumns.FlowDirection
        .LineBetween = s1.PageSetup.TextColumns.LineBetween
        If s1.PageSetup.TextColumns.Count > 1 Then
            For i = 1 To .Count
                .Item(i).Width = s1.PageSetup.TextColumns(i).Width
                If i < .Count Then
                    .Item(i).SpaceAfter = s1.PageSetup.TextColumns(i).SpaceAfter
                End If
            Next i
        End If
    End With
End Sub

Private Sub DuplicateBorderProperties(s1 As Section, s2 As Section)
    Dim i As Long
    For i = 1 To s2.Borders.Count
        With s2.Borders(i)
            .LineStyle = s1.Borders(i).LineStyle
            If .LineStyle <> wdLineStyleNone Then
                .LineWidth = s1.Borders(i).LineWidth
                .ArtStyle = s1.Borders(i).ArtStyle
                .ArtWidth = s1.Borders(i).ArtWidth
                .Color = s1.Borders(i).Color
                .Visible = s1.Borders(i).Visible
            End If
        End With
    Next i
    With s2.Borders
        .AlwaysInFront = s1.Borders.AlwaysInFront
        .DistanceFrom = s1.Borders.DistanceFrom
        .DistanceFromBottom = s1.Borders.DistanceFromBottom
        .DistanceFromLeft = s1.Borders.DistanceFromLeft
        .DistanceFromRight = s1.Borders.DistanceFromRight
        .DistanceFromTop = s1.Borders.DistanceFromTop
        '.Enable = s1.Borders.Enable ' Don't use - sets the line style to the default line style and sets the line width to the default line width. (See MSDN)
            ' Also see shaunakelly.com/word/layout/page-borders.html - setting applies to all sections
        .EnableFirstPageInSection = s1.Borders.EnableFirstPageInSection
        .EnableOtherPagesInSection = s1.Borders.EnableOtherPagesInSection
        '.JoinBorders = s1.Borders.JoinBorders ' apparent bug: removes borders from other sections
        '.SurroundFooter = s1.Borders.SurroundFooter ' apparent bug: removes borders from other sections
        '.SurroundHeader = s1.Borders.SurroundHeader  ' apparent bug: removes borders from other sections
    End With
End Sub

Private Sub DuplicateHeadersAndFooters(s1 As Section, s2 As Section)
    ' first link to previous (to copy them), then duplicate setting
    Dim i As Long
    For i = 1 To 3
        s2.Headers(i).LinkToPrevious = True
        s2.Headers(i).LinkToPrevious = s1.Headers(i).LinkToPrevious
        s2.Footers(i).LinkToPrevious = True
        s2.Footers(i).LinkToPrevious = s1.Footers(i).LinkToPrevious
    Next i
End Sub

Private Sub DuplicatePageNumbers(s1 As Section, s2 As Section)
    ' PageNumbers behaves like a property of the Section object, not a HeaderFooter object.
    ' If you change one property for one HeaderFooter.PageNumbers,
    ' it changes the same property for all other HeaderFooters.
    ' Therefore, only need to apply to one HeaderFooter object
    With s2.Footers(1).PageNumbers ' 1 is primary
        .NumberStyle = s1.Footers(1).PageNumbers.NumberStyle
        .RestartNumberingAtSection = s1.Footers(1).PageNumbers.RestartNumberingAtSection
        If .RestartNumberingAtSection Then
            .StartingNumber = s1.Footers(1).PageNumbers.StartingNumber
        End If
        If s1.Footers(1).PageNumbers.IncludeChapterNumber Then
            .IncludeChapterNumber = True
            .HeadingLevelForChapter = s1.Footers(1).PageNumbers.HeadingLevelForChapter
            .ChapterPageSeparator = s1.Footers(1).PageNumbers.ChapterPageSeparator
        Else
            .HeadingLevelForChapter = 0
            .IncludeChapterNumber = False
        End If
        .DoubleQuote = s1.Footers(1).PageNumbers.DoubleQuote
    End With
End Sub

