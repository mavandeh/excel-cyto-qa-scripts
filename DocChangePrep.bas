Attribute VB_Name = "DocChangePrep"
Private Function DocNumToURL(colnum As Long)
    
    Dim rng As Range, URL As String, ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(1)
    Set rng = Range(ws.Cells(2, colnum), ws.Cells(LastRow(ws), colnum))
    
    For Each cell In rng
        cell.Select
        If cell.Value <> "" Then
            If Left(cell.Value, 7) = "http://" Then
                URL = cell.Value
            Else
                URL = "http://mayoweb.mayo.edu/ap-docs/documents/0" & cell.Value & ".pdf"
            End If
        ActiveSheet.Hyperlinks.Add Anchor:=cell, Address:=URL
        End If
    Next
    
End Function

Private Function ImportReviewers(colnum As Long)
    
    ' make helper column and vlookup for reviewers
    Dim rng As Range, rng2 As Range, URL As String, ws As Worksheet
    Dim titlecol As Long, helpercol As Long, docidcol As Long, revcol As Long, lookupcol As Long, datecol As Long
        
    Set ws = ActiveWorkbook.Worksheets(1)
    
    
    ' find title, doc id, parties, helper and lookup columns
    For i = 1 To lastCol(ws)
        If Trim(Left(Cells(1, i).Value, 4)) = "Date" Then
            datecol = i
        ElseIf Trim(Cells(1, i).Value) = "Document ID" Then
            docidcol = i
        ElseIf Trim(Cells(1, i).Value) = "Title" Then
            titlecol = i
        ElseIf Trim(Cells(1, i).Value) = "Responsible Parties" Then
            revcol = i
        ElseIf Trim(Cells(1, i).Value) = "Helper Column" Then
            helpercol = i
        ElseIf Trim(Cells(1, i).Value) = "VLOOKUP Column" Then
            lookupcol = i
        End If
    Next
    
    ' set up helper columns with docid & date
    Set rng = Range(ws.Cells(2, helpercol), ws.Cells(LastRow(ws), helpercol))
    For Each cell In rng
        cell.Select
        cell.FormulaR1C1 = "=""0""&RC" & docidcol & "&RC" & datecol
    Next
            
    ' set up lookup column
    Set rng = Range(ws.Cells(2, lookupcol), ws.Cells(LastRow(ws), lookupcol))
    For Each cell In rng
        cell.Select
        cell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC" & helpercol & ",'[CytoMAP agenda.xlsx]2019'!C8:C9,2,FALSE),VLOOKUP(RC" & titlecol & ",'[CytoMAP agenda.xlsx]2019'!C4:C9,6,FALSE))"
    Next
    
    rng.Copy
    ws.Cells(2, revcol).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
        
End Function

Private Function lastCol(sh As Worksheet)
    ' Borrowed from https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
    On Error Resume Next
    lastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Function LastRow(sh As Worksheet)
    ' Borrowed from https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row
    On Error GoTo 0
End Function

Private Function DeleteOtherLabs(col As Long)
    
    Dim cyto As String, map As String, ap As String, i As Long, ws As Worksheet, row As Long
    
    cyto = "Cytopathology"
    map = "Molecular AP"
    ap = "Anatomic Pathology"
    Set ws = ActiveWorkbook.Worksheets(1)
    'start row at 2 so it can increment
    row = 2
    
    For i = 2 To LastRow(ws)
        
        If (Trim(ws.Cells(row, col).Value) = cyto) Or _
            (Trim(ws.Cells(row, col).Value) = map) Or _
            (Trim(ws.Cells(row, col).Value) = ap) Then
            'do nothing
        ElseIf ws.Cells(row, col).Value <> "" Or IsEmpty(ws.Cells(row, col)) Then
            ws.Rows(row).Delete
            row = row - 1
        End If
        row = row + 1
    Next i
    
End Function

Private Function AddSigLine(lrow As Long)
    
    ' OUTER BOX, BOLD PURPLE, ARIAL 14 BOLD ITALIC
    Dim oBox As Range, sigLine As Range, sigDate As Range, attestation As Range
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(1)
    Set oBox = ws.Range(ws.Cells(lrow + 1, 1), ws.Cells(lrow + 3, 9))
    Set attestation = ws.Cells(lrow + 1, 1)
    Set sigLine = ws.Range(ws.Cells(lrow + 2, 1), ws.Cells(lrow + 2, 6))
    Set sigDate = ws.Range(ws.Cells(lrow + 3, 1), ws.Cells(lrow + 3, 3))
    
    With oBox.Font
        .name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
        .Italic = True
        .Bold = True
    End With
    With oBox.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oBox.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oBox.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oBox.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With oBox.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    'signature line
    With sigLine
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    
    'add signature and date
    With sigDate
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With

    sigDate.Value = "Signature and Date"
    attestation.Value = "I have read and understand the contents of the listed documents"
    attestation.WrapText = False
    
End Function

Sub DocChangePrep()

    Dim urlcol As Long, labcol As Long, revcol As Long, ws As Worksheet, rng As Range
    Set ws = ActiveWorkbook.Worksheets(1)

    For urlcol = 1 To lastCol(ws)
        If Trim(ws.Cells(1, urlcol).Value) = "Document ID" Then
            Exit For
        End If
    Next urlcol
    For labcol = 1 To lastCol(ws)
        If LCase(Trim(ws.Cells(1, labcol).Value)) = "tier 3" Then
            Exit For
        End If
    Next labcol
    
    'If ws.Cells(1, LastCol(ws) + 1).Value <> "Initials & Date" Then
        ws.Cells(1, lastCol(ws)).Copy
        ws.Cells(1, lastCol(ws) + 1).Value = "Responsible Parties"
        ws.Cells(1, lastCol(ws)).PasteSpecial xlFormats
        ws.Cells(1, lastCol(ws) + 1).Value = "Initials & Date"
        ws.Cells(1, lastCol(ws)).PasteSpecial xlFormats
        ws.Cells(1, lastCol(ws) + 1).Value = "Helper Column"
        ws.Cells(1, lastCol(ws)).PasteSpecial xlFormats
        ws.Cells(1, lastCol(ws) + 1).Value = "VLOOKUP Column"
        ws.Cells(1, lastCol(ws)).PasteSpecial xlFormats
        revcol = lastCol(ws)
    'End If
    
    DeleteOtherLabs (labcol)
    DocNumToURL (urlcol)
    ImportReviewers (revcol)
        
    Set rng = Range(Cells(2, lastCol(ws) - 3), Cells(LastRow(ws), lastCol(ws)))
    
    rng.Borders.LineStyle = xlContinuous
    
    AddSigLine (LastRow(ws))
    
    For Each cell In Range(Cells(1, 1), Cells(1, lastCol(ws)))
        If cell.Value = "Helper Column" Or cell.Value = "VLOOKUP Column" Then
            cell.EntireColumn.Hidden = True
        End If
    Next

End Sub





