Attribute VB_Name = "PDxGYNInterps"
' This file is meant for cleaning up PathDx Cytology Multiple Interpretations Report (GYN Version)
' Author: Mark A. VandeHaar, SCT(ASCP)
'

Option Explicit

Private Sub UnmergeAll()
  Dim currentSheet As Worksheet
  For Each currentSheet In Worksheets
    currentSheet.Cells.unmerge
  Next
End Sub

Private Sub DeleteEmptyRows()
  Dim currentSheet As Worksheet
  For Each currentSheet In Worksheets
    On Error Resume Next
    currentSheet.Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
  Next
End Sub

Function LastRow(sh As Worksheet)
    ' Borrowed from https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(sh As Worksheet)
    ' Borrowed from https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

Private Sub CopyRangeFromMultiWorksheets()

    ' Modified from https://msdn.microsoft.com/en-us/library/cc793964(v=office.12).aspx
    Dim sh As Worksheet
    Dim DestSh As Worksheet
    Dim Last As Long
    Dim CopyRng As Range

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Delete the summary sheet if it exists.
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("Data").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Select Sheet1 because otherwise next statement will add several extra sheets
    ActiveWorkbook.Worksheets("Sheet1").Select
    
    ' Add a new summary worksheet.
    Set DestSh = ActiveWorkbook.Worksheets.Add
    DestSh.name = "Data"

    ' Loop through all worksheets and copy the data to the summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If sh.name <> DestSh.name Then

            ' Find the last row with data on the summary worksheet.
            Last = LastRow(DestSh)

            ' Specify the range to place the data.
            Set CopyRng = sh.UsedRange

            ' Test to see whether there are enough rows in the summary
            ' worksheet to copy all the data.
            If Last + CopyRng.Rows.Count > DestSh.Rows.Count Then
                MsgBox "There are not enough rows in the " & _
                   "summary worksheet to place the data."
                GoTo ExitTheSub
            End If

            ' This statement copies values and formats from each worksheet.
            CopyRng.Copy
            With DestSh.Cells(Last + 1, "A")
                .PasteSpecial xlPasteValues
                .PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End With

        End If
    Next
        
ExitTheSub:

    Application.GoTo DestSh.Cells(1)

    ' AutoFit the column width in the summary sheet.
    DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Private Sub HideSheets()
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
        If Left(sh.name, 5) = "Sheet" Then
            sh.Visible = xlSheetHidden
        End If
    Next sh
End Sub

Private Sub SortData()
    ' created from macro recording and modified
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("A1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("B1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "HPV,HPVG,TPRPS,TPRPD,STHPV,DTHPV,STPCO,DTPCO", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("P1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A:Z")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
End Sub

Private Sub UpdateHPVResults()
  
  With Application
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .ReferenceStyle = xlR1C1
  End With

    Dim HPV16Col As Integer, mankato As Boolean, mcra As Boolean, ws As Worksheet, i As Long
        
    Set ws = ActiveWorkbook.Worksheets("Data")
    'find the hpv16 column and store the index
    For HPV16Col = 1 To LastCol(ws)
        If ws.Cells(1, HPV16Col).Value = "HPVG1" Or ws.Cells(1, HPV16Col).Value = "HPV16" Then
            ws.Cells(1, HPV16Col).Value = "HPV16"
            Exit For
        End If
    Next HPV16Col
    Dim str As Variant
    'loop through all rows looking at accession classes to determine if data is mixed or MCHS or MCR/MCA only
    For i = 1 To LastRow(ws)
        If i > 1 Then ws.Cells(i, 1) = Left(ws.Cells(i, 1), 9)
        If (Left(ws.Cells(i, 1).Value, 2) = "GM") Then
            mankato = True
        ElseIf (Left(ws.Cells(i, 1).Value, 2) = "GA") Or (Left(ws.Cells(i, 1).Value, 2) = "GR") Then
            mcra = True
        End If
        If mcra And mankato Then
            'results are incompatible because test code is different.
            MsgBox "Sheet contains data with both GM and GR/GA cases.  HPV results are incompatible. " _
                & "Please rerun report including only filters for GM/HPVG or GR/GA/HPV and " _
                & "Pap test codes."
            End
        End If
    Next i

    'MsgBox "UpdateHPVResults(): mcra = " & mcra & ", mankato = " & mankato
    
    Dim sCustomList(1 To 9) As String
    sCustomList(1) = "HPV"
    sCustomList(2) = "HPVG"
    sCustomList(3) = "STPCO"
    sCustomList(4) = "DTPCO"
    sCustomList(5) = "STHPV"
    sCustomList(6) = "DTHPV"
    sCustomList(7) = "TPRPS"
    sCustomList(8) = "TPRPD"
    sCustomList(9) = "TPRCY"
    
    Application.AddCustomList ListArray:=sCustomList
    
    Dim Rngsort As Range, RngKey As Range, RngKey1 As Range
    
    'Populate Ws
    Set ws = ActiveWorkbook.Worksheets("Data")
    
    'Clear out any previous Sorts that may be leftover
    ws.Sort.SortFields.Clear
    
    'range that includes all columns to sort
    Set Rngsort = ws.UsedRange
    
    'Columns with keys to sort
    Set RngKey = ws.Range("A1")
    Set RngKey1 = ws.Range("B1")

    'Perform the sort
    With ActiveWorkbook.Worksheets("Data").Sort
        Rngsort.Sort Key1:=RngKey1, Order1:=xlAscending, Header:=xlYes, OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, Orientation:=xlSortColumns, DataOption1:=xlSortNormal
        Rngsort.Sort Key1:=RngKey, Order1:=xlAscending, Header:=xlYes, OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, Orientation:=xlSortColumns, DataOption1:=xlSortNormal
    End With
    
    Application.DeleteCustomList Application.CustomListCount

    Dim col As Integer
    For col = HPV16Col To (HPV16Col + 2)

      Dim cell As Range
      On Error Resume Next
      If Not mankato Then
        For Each cell In Columns(col).SpecialCells(xlCellTypeBlanks).Areas
            cell.FormulaArray = "=INDEX(C,MATCH(LEFT(RC1,9)&""HPV"",LEFT(C1,9)&C2,0))"
            cell.Value = cell.Value ' comment this line for formula troubleshooting
        Next cell
      ElseIf mankato Then
        If ws.Cells(1, col).Value = "HPVG18" Then ws.Cells(1, col).Value = "HPV18/45"
        For Each cell In Columns(col).SpecialCells(xlCellTypeBlanks).Areas
            cell.FormulaArray = "=INDEX(C,MATCH(LEFT(RC1,9)&""HPVG"",LEFT(C1,9)&C2,0))"
            cell.Value = cell.Value ' comment this line for formula troubleshooting
        Next cell
      End If
    Next col

    If Not mankato Then InsertHPVOverall
    
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
End Sub

Private Sub RenameSheet()
    ActiveSheet.name = "Data"
End Sub

Private Sub DeleteHPVLines()

    ' Modified from https://danwagner.co/how-to-delete-rows-with-range-autofilter/
    Dim ws As Worksheet, lr As Long, lc As Long
    Dim rngData As Range

    Set ws = ActiveWorkbook.Worksheets("Data")
    lr = LastRow(ws)
    lc = LastCol(ws)
    Set rngData = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc))
    
    
    Application.DisplayAlerts = False
        With rngData
            On Error Resume Next
            ' Filter for HPV in the test column (#2)
            .AutoFilter Field:=2, Criteria1:="HPV", Operator:=xlOr, Criteria2:="HPVG"
            ' Delete visible rows, keep header
            .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With
    Application.DisplayAlerts = True
    
    'Turn off the AutoFilter
    With ws
        .AutoFilterMode = False
        If .FilterMode = True Then
            .ShowAllData
        End If
    End With
  
End Sub

Private Sub DeleteDuplicateInterpretations()

  ' Original code utilizing LastRow() Function from MSDN above.

  Dim ws As Worksheet, lr As Long, lc As Long, empCol As Long, caseCol As Long, intDTCol As Long, i As Integer
  Dim str As String, uniqueID As String, spectypeCol As Long
                  
  Set ws = Worksheets("Data")
  lr = LastRow(ws)
  lc = LastCol(ws)
  uniqueID = "CASE_EMPLOYEE"            'this is the unique id by which duplicates will be identified,
                                        'concatenation of CASE NUMBER and EMPLOYEE
  
  'find case number and employee columns
  For i = 1 To lc
    'MsgBox "DeleteDuplicateInterpretations(): col(" & i & ") name is " & ws.Cells(1, i).Value
    If ws.Cells(1, i).Value = "CASE NUMBER" Then
        caseCol = i
    ElseIf ws.Cells(1, i).Value = "EMPLOYEE" Then
        empCol = i
    ElseIf ws.Cells(1, i).Value = "INTERPRETATION DT" Then
        intDTCol = i
    ElseIf ws.Cells(1, i).Value = "SPECIMEN TYPE" Then
        spectypeCol = i
    ElseIf (caseCol > 0) And (empCol > 0) And (intDTCol > 0) And (spectypeCol > 0) Then
        Exit For
    End If
    If i = lc Then
        MsgBox "DeleteDuplicateInterpretations(): Could not find one or more columns: " _
            & "case number, employee, specimen type, or interpretation dt."
        End
    End If
  Next i

  'Search backwards from last column to see if uniqueID column has been created.
  'if not, create uniqueID (case_employee) column
  For i = 0 To lc - 1
    If (ws.Cells(1, lc - i).Value = uniqueID) Then
        Exit For
    ElseIf i = lc - 1 Then
        ws.Cells(1, lc + 1).Value = uniqueID
        lc = LastCol(ws)
        ws.Cells(2, lc).Formula = "=Left(RC[" & caseCol - lc & "],9)&RC[" & spectypeCol - lc & "]&RC[" & empCol - lc & "]"
        ws.Cells(2, lc).AutoFill Destination:=Range(ws.Cells(2, lc), ws.Cells(lr, lc))
    End If
  Next i
  
  'sort by case-person ascending (Y) and then by interp date descending (P)
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ' sort on CASE_EMPLOYEE
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range(ws.Cells(2, lc), ws.Cells(lr, lc)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' sort on interpretation date and time
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range(ws.Cells(2, intDTCol), ws.Cells(lr, intDTCol)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range(ws.Cells(1, 1), ws.Cells(lr, lc))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
  'remove duplicates based on case-person column (will delete second entries leaving the most recent which is sorted at top)
  Range(ws.Cells(2, 1), ws.Cells(lr, lc)).RemoveDuplicates Columns:=Array(lc)
  
End Sub

Private Function CheckHPV(rng As String) As Boolean

  If Range(rng) = "HPVG1" Or Range(rng) = "HPV16" Then
  CheckHPV = True
  Else
  CheckHPV = False
  End If
  
End Function

Private Sub InsertHPVOverall()

    'inserthpvoverall() is only run if there are no mankato cases in the spreadsheet.
    With Application
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .ReferenceStyle = xlR1C1
    End With
    
    Dim ws As Worksheet, lr As Long, lc As Long, colName As String, i As Integer
    Dim HPV16Col As Long, HPV18Col As Long, HPVOthCol As Long
            
    colName = "HPVOverall"                      'if this name changes, it must also change in PTHPVbyDx and PTASCUSHPV
    Set ws = ActiveWorkbook.Worksheets("Data")
    lr = LastRow(ws)
    lc = LastCol(ws)
    
    'search columns backwards for if overall hpv column already exists, if so, delete
    For i = 0 To lc - 1
        If ws.Cells(1, lc - i).Value = colName Then
            ws.Range(ws.Cells(1, lc - i), ws.Cells(1, lc - i)).EntireColumn.Delete
            lc = LastCol(ws)
            Exit For
        End If
    Next i
        
    ' column location invariant overall placement
    For i = 1 To lc
        If ws.Cells(1, i).Value = "HPVG1" Or ws.Cells(1, i).Value = "HPV16" Then
            HPV16Col = i
        ElseIf ws.Cells(1, i).Value = "HPVG18" Or ws.Cells(1, i).Value = "HPV18" Then
            HPV18Col = i
        ElseIf ws.Cells(1, i).Value = "HPVGOTHER" Or ws.Cells(1, i).Value = "HPVOTHER" Then
            HPVOthCol = i
        ElseIf (HPV16Col > 0) And (HPV18Col > 0) And (HPVOthCol > 0) Then
            Exit For
        ElseIf i = lc Then
            MsgBox "InsertHPVOverall(): Could not find one or more columns: " _
                & "HPVG1/HPV16, HPV(G)18, or HPV(G)OTHER."
            End
        End If
    Next i
    
    ws.Cells(1, lc + 1).Value = colName
    Range(ws.Cells(2, lc + 1), ws.Cells(lr, lc + 1)).FormulaR1C1 = "=IF(OR(RC" & HPV16Col & "=""Positive"",OR(RC" & HPV18Col & "=""Positive"",RC" & HPVOthCol & "=""Positive"")),""Positive"",IF(OR(RC" & HPV16Col & "=""Negative"",OR(RC" & HPV18Col & "=""Negative"",RC" & HPVOthCol & "=""Negative"")),""Negative"",0))"
    
    With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
  End With

End Sub

Private Sub RowSizeZoom()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    ws.Range("A2:A" & ws.Rows.Count).rowHeight = 12.75
    ws.Activate
    ActiveWindow.Zoom = 85
  Next ws
  
End Sub

Sub CleanUp()
    Dim scount As Long
    scount = ThisWorkbook.Sheets.Count
    If scount > 1 Then
        ' MsgBox "this workbook has multiple sheets"
        MultiSheetSub
    Else
    If scount = 1 Then
        ' MsgBox "this workbook has only ONE sheet"
        SingleSheetSub
        End If
    End If
            
End Sub

Private Sub PTInterpTotals()
'
' PTInterpTotals Macro
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, lr As Long, pt As PivotTable, name As String, ws As Worksheet
    Dim fArr, rngList As Range, elem As Variant, i As Long
    Dim pi As PivotItem, pf As PivotField
           
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets("InterpTotals").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    name = "InterpTotals"
    
    
    Worksheets("Data").Activate
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(192, 0, 0)
    End With
    
    Set ws = ActiveWorkbook.Worksheets(name)
    
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:="PT" & name)

    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 4
    End With

    With pt.PivotFields("QUALITY CODE")
        .Orientation = xlColumnField
        .Position = 1
    End With

    'Sort quality codes in ascending order on chart.
    With pt.PivotFields("QUALITY CODE")
        On Error Resume Next
        .PivotItems("CYAGREE").Position = 1
        .PivotItems("CYMINOR").Position = 2
        .PivotItems("CYMAJOR").Position = 3
        .PivotItems("CY+3").Position = 1
        .PivotItems("CY+2").Position = 1
        .PivotItems("CY+1.5").Position = 1
        .PivotItems("CY+1").Position = 1
        .PivotItems("CY+0.5").Position = 1
        .PivotItems("CY0").Position = 1
        .PivotItems("CY-0.5").Position = 1
        .PivotItems("CY-1").Position = 1
        .PivotItems("CY-1.5").Position = 1
        .PivotItems("CY-2").Position = 1
        .PivotItems("CY-3").Position = 1
    End With

    'select discrepancies and group them:
    'if pivot item is there, add it to array, then loop through the array, select and group
    
    'create filter list
    rngList = ActiveWorkbook.Worksheets(name).Range("BB1:BB12")
    fArr = Array("CY-3", "CY-2", "CY-1.5", "CY-1", "CY-0.5", "CY+0.5", "CY+1", "CY+1.5", "CY+2", "CY+3", "CYMINOR", "CYMAJOR")
    i = 1
    
    'adds quality codes to a range for filtering?
    For Each elem In fArr
        ws.Range("BB" & i) = elem
        i = i + 1
    Next elem
    
    Set pf = pt.PivotFields("QUALITY CODE")
    

    Dim val As String
    
    'filter and group discrepancies
    For Each pi In pf.PivotItems
        val = pi.Value
        If (val = "CY+3") Or (val = "CY+2") Or (val = "CY+1.5") Or (val = "CY+1") Or (val = "CY+0.5") _
            Or (val = "CY-3") Or (val = "CY-2") Or (val = "CY-1.5") Or (val = "CY-1") Or (val = "CY-0.5") _
            Or (val = "CYMINOR") Or (val = "CYMAJOR") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    Application.PivotTableSelection = True
    pt.PivotSelect "QUALITY CODE[All]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "QUALITY CODE2[Group1]", xlLabelOnly
    Selection.Value = "DISCREPANCIES"
    
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    
    Set pf = pt.PivotFields("QUALITY CODE2")
    
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    
    'filter and group agrees
    For Each pi In pf.PivotItems
        val = pi.Value
        If (val = "CY0") Or (val = "CYAGREE") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    Application.PivotTableSelection = True
    pt.PivotSelect "QUALITY CODE[All]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "QUALITY CODE2[Group2]", xlLabelOnly
    Selection.Value = "AGREES"
    
    Set pf = pt.PivotFields("QUALITY CODE")
    
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    
    Set pf = pt.PivotFields("QUALITY CODE2")
    
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    
    'collapse to employee
    pt.PivotFields("EMPLOYEE").ShowDetail = False
    pt.PivotFields("QUALITY CODE2").ShowDetail = False
    
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    pt.PivotFields("EMPLOYEE").ShowDetail = False
    
    Range("J1").Value = "To send discrepant case list by employee:"
    Range("K2").Value = "Double click on each employee's DISCREPANCIES number to generate new sheet."
    Range("K3").Value = "Right click on sheet tab and select move, select (new book) from dropdown."
    Range("K4").Value = "In new book go to File > Share > Email > Send as Attachment."
    Range("K4").Value = "Enter the employee's email, and click send."
    
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False

    
End Sub

Private Sub zASCtoSIL()

    Dim ws As Worksheet, pt As PivotTable, pi As PivotItem, pt2 As PivotTable
    Dim i As Long, j As Long, empRng As Range
    Dim pasteCol As Long, s As String, lcol As Long, lrow As Long
    Dim auCol As Long, ahCol As Long, lsCol As Long, hsCol As Long
    
    Set ws = Worksheets("Benchmarks")
    Set pt = ws.PivotTables("PTBenchmarksCount")
    Set pt2 = ws.PivotTables("PTBenchmarksPercent")
    pasteCol = LastCol(ws) + 2
    j = 1
        
    With [A1]
        pt.TableRange2.Copy
        ws.Cells(1, pasteCol).PasteSpecial xlPasteValues
    End With
    
    For i = pasteCol To pasteCol + 15
        s = Cells(2, i).Value
        If (s <> "GYN ASCUS") And (s <> "GYN ASCH") And (s <> "GYN LSIL") And (s <> "GYN HSIL") And _
          (s <> "") Then
            Cells(2, i).EntireColumn.Hidden = True
        ElseIf (s <> "") Then
            If s = "GYN ASCUS" Then
                auCol = i
            ElseIf s = "GYN ASCH" Then
                ahCol = i
            ElseIf s = "GYN LSIL" Then
                lsCol = i
            ElseIf s = "GYN HSIL" Then
                hsCol = i
            End If

        End If
        
    Next i
    
    For i = 4 To LastRow(ws)
        If Len(ws.Cells(i, pasteCol).Value) > 5 Then
            With Range(Cells(i, pasteCol), Cells(i, LastCol(ws)))

                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
              
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlMedium
            
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlMedium
                
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlMedium
                
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlMedium
    

                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                .Font.Bold = True
            End With
        End If
    Next i
    
    pasteCol = LastCol(ws) + 2
    
    'create static percentage table for benchmarks
    With [A1]
        pt2.TableRange2.Copy
        ws.Cells(1, pasteCol).PasteSpecial xlPasteValues
    End With
    
    lcol = LastCol(ws)
    lrow = LastRow(ws)
    Cells(2, lcol + 1).Value = "ASC:SIL Ratio"
    
    For i = 4 To lrow
        On Error Resume Next
        Cells(i, lcol + 1).Value = (Cells(i, auCol).Value + Cells(i, ahCol).Value) _
            / (Cells(i, lsCol).Value + Cells(i, hsCol).Value)
    Next i
    
    For i = 4 To LastRow(ws)
        If Len(ws.Cells(i, pasteCol).Value) > 5 Then
            With Range(Cells(i, pasteCol), Cells(i, LastCol(ws)))

                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
              
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).ColorIndex = 0
                .Borders(xlEdgeLeft).TintAndShade = 0
                .Borders(xlEdgeLeft).Weight = xlMedium
            
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).ColorIndex = 0
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlMedium
                
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).ColorIndex = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).Weight = xlMedium
                
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).ColorIndex = 0
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeRight).Weight = xlMedium
    

                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                .Font.Bold = True
            End With
        End If
    Next i

End Sub

Private Sub PTBenchmarks()
'
' PTBenchmarks Macro
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, pt As PivotTable, pi As PivotItem, pt2 As PivotTable
    Dim ws As Worksheet, name As String, lcol As Long, lrow As Long, i As Integer, j As Integer
    
        
    name = "Benchmarks"         'worksheet name, used to construct pivot table names
    
    
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Worksheets("Data").Activate
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(237, 125, 49)
    End With
    
    Set ws = Sheets(name)
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:="PT" & name & "Count")

    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 4
    End With
    
    'Diagnosis Category Setup
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    pt.PivotFields("EMPLOYEE").ShowDetail = False

    'select nil category items and group them:
    On Error Resume Next
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    
    'select agus category items and group them
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN AGUS,GYN AIS]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group2]", xlLabelOnly
    Selection.Value = "AGUS"
    
    'sort diagnosis categories
    With pt.PivotFields("DIAGNOSIS CATEGORY2")
        On Error Resume Next
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("AGUS").Position = 7
        .PivotItems("GYN CANCER").Position = 8
    End With
    
    'hide blanks
    pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
    
    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    lcol = LastCol(ws) + 2
    Set pt2 = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=ws.Cells(1, lcol), TableName:="PTBenchmarksPercent")
    
    With pt2.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt2.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt2.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt2.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 4
    End With
    'collapse to employee
    pt2.PivotFields("EMPLOYEE").ShowDetail = False
    
    'Diagnosis Category Setup
    With pt2.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt2.PivotFields("DIAGNOSIS CATEGORY2")
        .Orientation = xlColumnField
        .Position = 1
    End With

    'select nil category items and group them:
    Application.PivotTableSelection = True
'    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
'    Selection.Group
    pt2.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    pt2.PivotSelect "DIAGNOSIS CATEGORY2[Group2]", xlLabelOnly
    Selection.Value = "AGUS"
    
    'sort diagnosis categories
    With pt2.PivotFields("DIAGNOSIS CATEGORY2")
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("AGUS").Position = 7
        .PivotItems("GYN CANCER").Position = 8
    End With
    
    'hide blanks
    pt2.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
    
    'add interp count by case number and collapse to employee
    pt2.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    'change second PT to percentage
    With pt2.PivotFields("Count of CASE NUMBER")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0.00%"
    End With
    
    zASCtoSIL
    
    
    
    'populate formatting lookup table.  Numbers from 2016 benchmarks.
    
    'Table and category titles
    'prepare to populate benchmarks table.
    Dim bMarksTitleCol As Long
    lcol = LastCol(ws) + 2
    'store value of this column for later
    bMarksTitleCol = lcol
    With ws
        .Cells(1, lcol).FormulaR1C1 = "CAP Benchmarks"
        .Cells(2, lcol).FormulaR1C1 = "Diagnosis Category"
        .Cells(4, lcol).FormulaR1C1 = "GYN ASCUS"
        .Cells(5, lcol).FormulaR1C1 = "GYN ASCH"
        .Cells(6, lcol).FormulaR1C1 = "GYN LSIL"
        .Cells(7, lcol).FormulaR1C1 = "GYN HSIL"
        .Cells(8, lcol).FormulaR1C1 = "AGUS"
        .Cells(9, lcol).FormulaR1C1 = "GYN UNSAT"
        .Cells(10, lcol).FormulaR1C1 = "ASC:SIL Ratio"
    End With
    
    '5th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "5th"
        .Cells(4, lcol).FormulaR1C1 = "0.021"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0"           'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.011"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.001"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0"           'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.003"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "0.8"        'ASC:SIL
    End With
    
    '10th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "10th"
        .Cells(4, lcol).FormulaR1C1 = "0.027"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.001"       'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.014"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.002"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0"           'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.004"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "0.9"        'ASC:SIL
    End With
    
    '25th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "25th"
        .Cells(4, lcol).FormulaR1C1 = "0.039"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.002"       'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.02"        'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.003"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0.001"       'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.008"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "1.4"        'ASC:SIL
    End With

    '50th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "50th"
        .Cells(4, lcol).FormulaR1C1 = "0.054"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.003"       'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.027"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.004"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0.002"       'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.013"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "1.8"        'ASC:SIL
    End With
    
    '75th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "75th"
        .Cells(4, lcol).FormulaR1C1 = "0.075"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.005"       'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.036"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.007"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0.003"       'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.021"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "2.5"        'ASC:SIL
    End With
    
    '90th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "90th"
        .Cells(4, lcol).FormulaR1C1 = "0.103"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.008"       'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.047"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.011"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0.005"       'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.034"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "3.2"        'ASC:SIL
    End With

    '95th percentile values
    lcol = LastCol(ws) + 1
    With ws
        .Cells(2, lcol).FormulaR1C1 = "95th"
        .Cells(4, lcol).FormulaR1C1 = "0.125"       'ASCUS
        .Cells(5, lcol).FormulaR1C1 = "0.01"        'ASCH
        .Cells(6, lcol).FormulaR1C1 = "0.055"       'LSIL
        .Cells(7, lcol).FormulaR1C1 = "0.014"       'HSIL
        .Cells(8, lcol).FormulaR1C1 = "0.009"       'AGUS
        .Cells(9, lcol).FormulaR1C1 = "0.043"       'UNSAT
        .Cells(10, lcol).FormulaR1C1 = "3.8"        'ASC:SIL
    End With

    'set conditional formatting for static percentile columns
    Dim matchCount As Long
    matchCount = 0
    For i = 0 To 19
        For j = 4 To 10
            If Cells(2, lcol - i).Value = Cells(j, bMarksTitleCol).Value Then

                With Columns(lcol - i)
                    .FormatConditions.AddColorScale ColorScaleType:=3
                    '.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                'set lowest value
                    .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueFormula
                    .FormatConditions(1).ColorScaleCriteria(1).Value = "=Benchmarks!R" & j & "C" & lcol - 6
                    .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = 13011546
                    .FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0
                'set middle value
                    .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueFormula
                    .FormatConditions(1).ColorScaleCriteria(2).Value = "=Benchmarks!R" & j & "C" & lcol - 3
                    .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = 16776444
                    .FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = 0
                'set highest value
                    .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueFormula
                    .FormatConditions(1).ColorScaleCriteria(3).Value = "=Benchmarks!R" & j & "C" & lcol
                    .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = 7039480
                    .FormatConditions(1).ColorScaleCriteria(3).FormatColor.TintAndShade = 0
                End With
                matchCount = matchCount + 1
            End If
        Next j
        If matchCount = 7 Then Exit For
    Next i
    
    Dim fcolp As Long
    fcolp = 0
    While IsEmpty(Cells(3, LastCol(ws) - fcolp))
        fcolp = fcolp + 1
    Wend

    fcolp = LastCol(ws) - fcolp

    For i = 5 To LastRow(ws)
        If Len(ws.Cells(i, fcolp).Value) > 5 Then
            Range(Cells(i, fcolp), Cells(LastRow(ws), LastCol(ws) - 10)).FormatConditions.Delete
            Exit For
        End If
    Next i
    
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
End Sub

Private Sub PTHPVbyDx()
'
' PTHPVbyDx Macro
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, lr As Long, pt As PivotTable, pi As PivotItem, ws As Worksheet
    Dim name As String, ptName As String, cTitle As String, cLoc As String, i As Long, hpvColName As String
        
    name = "HPVbyDx"
    ptName = "PT" & name
    cTitle = "HPV Results by Diagnosis"
    cLoc = "E1"
        
    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Worksheets("Data").Activate
    Set ws = ActiveSheet
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(255, 192, 0)
    End With
    
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:=ptName)

    'loop through all rows looking at accession classes to determine if data is mixed or MCHS or MCR/MCA only

    Dim mankato As Boolean, mcra As Boolean
    For i = 1 To LastRow(ws)
        If i > 1 Then ws.Cells(i, 1) = Left(ws.Cells(i, 1), 9)
        If (Left(ws.Cells(i, 1).Value, 2) = "GM") Then
            mankato = True
            hpvColName = "HPVGOTHER"
        ElseIf (Left(ws.Cells(i, 1).Value, 2) = "GA") Or (Left(ws.Cells(i, 1).Value, 2) = "GR") Then
            mcra = True
            hpvColName = "HPVOVERALL"
        End If
        If mcra And mankato Then
            'results are incompatible because test code is different.
            MsgBox "Sheet contains data with both GM and GR/GA cases.  HPV results are incompatible. " _
                & "Please rerun report including only filters for GM/HPVG or GR/GA/HPV and " _
                & "Pap test codes."
            End
        End If
    Next i

    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    'select nil category items and group them:
    On Error Resume Next
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    'sort diagnosis categories
    With pt.PivotFields("DIAGNOSIS CATEGORY2")
        On Error Resume Next
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("GYN AGUS").Position = 7
        .PivotItems("GYN AIS").Position = 8
        .PivotItems("GYN CANCER").Position = 9
    End With
    
    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 4
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    With pt.PivotFields(hpvColName)
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("Positive").Position = 1
        .PivotItems("Negative").Position = 2
    End With
    With pt.PivotFields("TEST CODE")
        .Orientation = xlPageField
        .Position = 1
    End With

    For Each pi In pt.PivotFields(hpvColName).PivotItems
        If (pi.Value = "Positive") Or (pi.Value = "Negative") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi

    pt.PivotFields("EMPLOYEE").ShowDetail = False
    
    'hide blanks
    pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
    
    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    pt.PivotSelect ("")
    Charts.Add
    
    ActiveChart.Location Where:=xlLocationAsObject, name:=pt.Parent.name
    ActiveChart.Parent.Left = Range(cLoc).Left
    ActiveChart.Parent.Top = Range(cLoc).Top
    ActiveChart.ApplyLayout (3)
    ActiveChart.ChartType = xlColumnStacked100
    ActiveChart.ShowAllFieldButtons = False
    ActiveChart.HasTitle = True
    ActiveChart.chartTitle.Text = cTitle
   
    Selection.Format.TextFrame2.TextRange.Characters.Text = cTitle

    With ActiveChart.Parent
        .Height = 600 ' resize 2.5 pt at 72 ppi.
        .Width = 1000 ' resize 4.0 pt at 72 ppi.
    End With
    
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SeriesCollection(2).ApplyDataLabels
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    'add slicer buttons
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "TEST CODE").Slicers.Add ActiveSheet, , "TEST CODE", "TEST CODE", 238.5, 709.5 _
        , 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "EMPLOYEE").Slicers.Add ActiveSheet, , "EMPLOYEE 1", "EMPLOYEE", 276, 747, 144 _
        , 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "EMPLOYEE TYPE").Slicers.Add ActiveSheet, , "EMPLOYEE TYPE", "EMPLOYEE TYPE", _
        313.5, 784.5, 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "DIAGNOSIS CATEGORY").Slicers.Add ActiveSheet, , "DIAGNOSIS CATEGORY 1", _
        "DIAGNOSIS CATEGORY", 351, 822, 144, 187.5
    
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("F50").Value = "To filter by individuals, right click on their SoftID, Filter > Keep Only Selected Items, then drill down to DIAGNOSIS CATEGORY level."
    
End Sub

Private Sub PTASCUSHPV()
'
' PTASCUSHPV
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, lr As Long, pt As PivotTable, pi As PivotItem, ws As Worksheet
    Dim name As String, ptName As String, cTitle As String, cLoc As String, i As Long, hpvColName As String
        
    name = "ASCUSHPV"                                       'Pivot table and tab name
    ptName = "PT" & name
    cTitle = "HPV Results for ASCUS Cases by Pathologist"   'chart title
    cLoc = "E1"

    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Worksheets("Data").Activate
    Set ws = ActiveSheet
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(112, 173, 71)
    End With

    
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:=ptName)
    
    'loop through all rows looking at accession classes to determine if data is mixed or MCHS or MCR/MCA only
    Dim mankato As Boolean, mcra As Boolean
    For i = 1 To LastRow(ws)
        If i > 1 Then ws.Cells(i, 1) = Left(ws.Cells(i, 1), 9)
        If (Left(ws.Cells(i, 1).Value, 2) = "GM") Then
            mankato = True
            hpvColName = "HPVGOTHER"
        ElseIf (Left(ws.Cells(i, 1).Value, 2) = "GA") Or (Left(ws.Cells(i, 1).Value, 2) = "GR") Then
            mcra = True
            hpvColName = "HPVOVERALL"
        End If
        If mcra And mankato Then
            'results are incompatible because test code is different.
            MsgBox "Sheet contains data with both GM and GR/GA cases.  HPV results are incompatible. " _
                & "Please rerun report including only filters for GM/HPVG or GR/GA/HPV and " _
                & "Pap test codes."
            End
        End If
    Next i

    'Diagnosis category setup.  1) Add diagnosis category to PT
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    '2) select nil category items and group them:
    On Error Resume Next
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    
    '3) sort diagnosis categories
    With pt.PivotFields("DIAGNOSIS CATEGORY2")
        On Error Resume Next
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("GYN AGUS").Position = 7
        .PivotItems("GYN AIS").Position = 8
        .PivotItems("GYN CANCER").Position = 9
    End With
    
    'Add column/row fields and filters
    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 4
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    With pt.PivotFields(hpvColName)
        .Orientation = xlColumnField
        .Position = 1
        .PivotItems("Positive").Position = 1
        .PivotItems("Negative").Position = 2
    End With
    With pt.PivotFields("TEST CODE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    'Filter out all HPV results other than Positive and Negative
    For Each pi In pt.PivotFields(hpvColName).PivotItems
        If (pi.Value = "Positive") Or (pi.Value = "Negative") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    'Filter out all cytotechs.  These can be added back in by user
    For Each pi In pt.PivotFields("EMPLOYEE TYPE").PivotItems
        If (pi.Value = "Cytotechnologist") Or (pi.Value = "Technologist") _
            Or (pi.Value = "Lead Technologist") Then
            pi.Visible = False
        Else: pi.Visible = True
        End If
    Next pi

    pt.PivotFields("EMPLOYEE").ShowDetail = False
    
    'hide blanks and diagnoses that are not ascus
    For Each pi In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (pi.Value = "GYN ASCUS") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    pt.PivotSelect ("")
    Charts.Add
    
    ActiveChart.Location Where:=xlLocationAsObject, name:=pt.Parent.name
    ActiveChart.Parent.Left = Range(cLoc).Left
    ActiveChart.Parent.Top = Range(cLoc).Top
    ActiveChart.ApplyLayout (3)
    ActiveChart.ChartType = xlColumnStacked100
    ActiveChart.ShowAllFieldButtons = False
    ActiveChart.HasTitle = True
    ActiveChart.chartTitle.Text = cTitle
   
    Selection.Format.TextFrame2.TextRange.Characters.Text = cTitle

    With ActiveChart.Parent
        .Height = 600 ' resize 2.5 pt at 72 ppi.
        .Width = 1000 ' resize 4.0 pt at 72 ppi.
    End With
    
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SeriesCollection(2).ApplyDataLabels
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    'add slicer buttons
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PTASCUSHPV"), _
        "EMPLOYEE").Slicers.Add ActiveSheet, , "EMPLOYEE 2", "EMPLOYEE", 216, 621, 144 _
        , 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PTASCUSHPV"), _
        "EMPLOYEE TYPE").Slicers.Add ActiveSheet, , "EMPLOYEE TYPE 1", "EMPLOYEE TYPE" _
        , 253.5, 658.5, 144, 187.5
    
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("F50").Value = "Chart includes only ASCUS cases. Cytotechnologists may be added back in by clicking the filter icon to the right of EMPLOYEE TYPE in the Pivot Table Fields list (Select Pivot Table, Analyze Tab > Field List to show)."
    
End Sub

Private Sub PTCTAgreement()
'
' PTCTAgreement
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, lr As Long, pt As PivotTable, pi As PivotItem
    Dim name As String, ptName As String, cTitle As String, cLoc As String, i As Long
    
    name = "CTAgreement"                                       'Pivot table and tab name
    ptName = "PT" & name
    cTitle = "Cytotech-Cytotech Agreement Rate"                'chart title
    cLoc = "F1"                                                'cell for chart location

    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Worksheets("Data").Activate
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(91, 155, 213)
    End With
    
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:=ptName)

    'Diagnosis category setup.  1) Add diagnosis category to PT
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    '2) select nil category items and group them:
    On Error Resume Next
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    
    '3) sort diagnosis categories
    With pt.PivotFields("DIAGNOSIS CATEGORY2")
        On Error Resume Next
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("GYN AGUS").Position = 7
        .PivotItems("GYN AIS").Position = 8
        .PivotItems("GYN CANCER").Position = 9
        .Orientation = xlHidden
    End With
    
    
    'Add column/row fields and filters
    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlRowField
        .Position = 3 '4 because of diagnosis category setup will place it at 3.  This should go after.
    End With
    With pt.PivotFields("FINAL DIAGNOSIS")
        .Orientation = xlRowField
        .Position = 4 '4 because of diagnosis category setup will place it at 3.  This should go after.
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 7
    End With
    With pt.PivotFields("QUALITY CODE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt.PivotFields("TEST CODE")
        .Orientation = xlPageField
        .Position = 1
    End With
    
    'Sort quality codes in ascending order on chart.
    With pt.PivotFields("QUALITY CODE")
        .PivotItems("CYAGREE").Position = 1
        .PivotItems("CYMINOR").Position = 2
        .PivotItems("CYMAJOR").Position = 3
    End With
    
    'Filter out all HPV results other than Positive and Negative
    For Each pi In pt.PivotFields("QUALITY CODE").PivotItems
        If (pi.Value = "CYAGREE") Or (pi.Value = "CYMINOR") _
            Or (pi.Value = "CYMAJOR") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    'Filter out all cytotechs.  These can be added back in by user
    For Each pi In pt.PivotFields("EMPLOYEE TYPE").PivotItems
        If (pi.Value = "Cytotechnologist") Or (pi.Value = "Technologist") _
            Or (pi.Value = "Lead Technologist") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    'hide GYN UNSAT and other diagnoses irrelevant to rescreen
    For Each pi In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (pi.Value = "GYN NIL") Or (pi.Value = "GYNNOEC") Or (pi.Value = "GYN ORG") Then
            pi.Visible = True
        Else: pi.Visible = False
        End If
    Next pi
    
    pt.PivotFields("CASE NUMBER").ShowDetail = False
    pt.PivotFields("FINAL DIAGNOSIS").ShowDetail = False
    pt.PivotFields("EMPLOYEE").ShowDetail = False
    
    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    
    pt.PivotSelect ("")
    Charts.Add
    
    ActiveChart.Location Where:=xlLocationAsObject, name:=pt.Parent.name
    ActiveChart.Parent.Left = Range(cLoc).Left
    ActiveChart.Parent.Top = Range(cLoc).Top
    ActiveChart.ApplyLayout (3)
    ActiveChart.ChartType = xlColumnStacked100
    ActiveChart.ShowAllFieldButtons = False
    ActiveChart.HasTitle = True
    ActiveChart.chartTitle.Text = cTitle
   
    Selection.Format.TextFrame2.TextRange.Characters.Text = cTitle

    ActiveChart.SeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SeriesCollection(3).ApplyDataLabels
    ActiveChart.SeriesCollection(2).ApplyDataLabels
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    
    ActiveChart.ChartArea.Interior.Color = RGB(255, 255, 255)

    
    With ActiveChart.Parent
        .Height = 600 ' resize 2.5 pt at 72 ppi.
        .Width = 1000 ' resize 4.0 pt at 72 ppi.
    End With
    
    'add slicer buttons
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "EMPLOYEE").Slicers.Add ActiveSheet, , "EMPLOYEE", "EMPLOYEE", 216, 621, 144, _
        187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables(ptName), _
        "DIAGNOSIS CATEGORY").Slicers.Add ActiveSheet, , "DIAGNOSIS CATEGORY", _
        "DIAGNOSIS CATEGORY", 253.5, 658.5, 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveChart.PivotLayout.PivotTable, _
        "QUALITY CODE").Slicers.Add ActiveSheet, , "QUALITY CODE", "QUALITY CODE", _
        319.5, 696, 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveChart.PivotLayout.PivotTable, _
        "QUALITY REASON").Slicers.Add ActiveSheet, , "QUALITY REASON", "QUALITY REASON" _
        , 385.5, 733.5, 144, 187.5
    
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("F50").Value = "MCR: Quality Reason (blanks) represent late HPV entries and some training cases"
    Range("F51").Value = "All Sites: Drill down to look at diagnostic categories (what was changed to what). Expanding once displays Cytotech DxCategory followed by the nested list of what those cases were signed out as."
    Range("F52").Value = "All Sites: If a diagnosis is changed by the original cytotech before rescreen, the QA code is generated off of comparison to the first.  This means that you may see in this report CYMINOR on a case that was called NIL and changed to NIL.  The cytotech may have initially called it NOEC."
    Range("G53").Value = "To view this change in greater detail, go into processing history and click the ellipsis to the right of the diagnosis change."
    
End Sub

Private Sub PTCTPathAgreement()
'
' PTCTPathAgreement
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, lr As Long, pt As PivotTable, pi As PivotItem, ws As Worksheet, i As Long, j As Long
    Dim name As String, ptName As String, cTitle As String, cLoc As Range, pasteCol As Long, pasteRow As Long
    Dim chartatlastcol As Boolean
            
    name = "CTPathAgreement"                                   'Pivot table and tab name
    ptName = "PT" & name
    cTitle = "Cytotech-Pathologist Agreement Rate"             'chart title
    Set cLoc = Range("M1")                                     'cell for chart location
    chartatlastcol = True
    

    On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Worksheets("Data").Activate
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Range("A1").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(112, 48, 160)
    End With
    
    Set ws = Worksheets("CTPathAgreement")
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:=ptName)

    'Diagnosis category setup.  1) Add diagnosis category to PT
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    '2) select nil category items and group them:
    On Error Resume Next
    Application.PivotTableSelection = True
    pt.PivotSelect "DIAGNOSIS CATEGORY[GYN NIL,GYNNOEC,GYN REAC,GYN ORG]", xlLabelOnly
    Selection.Group
    pt.PivotSelect "DIAGNOSIS CATEGORY2[Group1]", xlLabelOnly
    Selection.Value = "NIL"
    
    '3) sort diagnosis categories
    With pt.PivotFields("DIAGNOSIS CATEGORY2")
        On Error Resume Next
        .ShowDetail = False
        .PivotItems("GYN UNSAT").Position = 1
        .PivotItems("NIL").Position = 2
        .PivotItems("GYN ASCUS").Position = 3
        .PivotItems("GYN ASCH").Position = 4
        .PivotItems("GYN LSIL").Position = 5
        .PivotItems("GYN HSIL").Position = 6
        .PivotItems("GYN AGUS").Position = 7
        .PivotItems("GYN AIS").Position = 8
        .PivotItems("GYN CANCER").Position = 9
    End With
    
    'Add column/row fields and filters
    With pt.PivotFields("EMPLOYEE TYPE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("EMPLOYEE")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("FINAL DIAGNOSIS")
        .Orientation = xlRowField
        .Position = 4 '4 because of diagnosis category setup will place it at 3.  This should go after.
    End With
    With pt.PivotFields("INTERPRETATION DT")
        .Orientation = xlRowField
        .Position = 5
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 6
    End With
    With pt.PivotFields("QUALITY CODE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt.PivotFields("TEST CODE")
        .Orientation = xlPageField
        .Position = 1
    End With

    'Filter out all HPV results other than Positive and Negative
    For Each pi In pt.PivotFields("QUALITY CODE").PivotItems
        If (pi.Value = "CYAGREE") Or (pi.Value = "CYMINOR") _
            Or (pi.Value = "CYMAJOR") Or (pi.Value = "(blank)") _
            Or (pi.Value = "PDSNOMED") Then
            pi.Visible = False
        Else: pi.Visible = True
        End If
    Next pi
    
    
    'Sort quality codes in ascending order on chart.
    With pt.PivotFields("QUALITY CODE")
        On Error Resume Next
        .PivotItems("CY+3").Position = 1
        .PivotItems("CY+2").Position = 1
        .PivotItems("CY+1.5").Position = 1
        .PivotItems("CY+1").Position = 1
        .PivotItems("CY+0.5").Position = 1
        .PivotItems("CY0").Position = 1
        .PivotItems("CY-0.5").Position = 1
        .PivotItems("CY-1").Position = 1
        .PivotItems("CY-1.5").Position = 1
        .PivotItems("CY-2").Position = 1
        .PivotItems("CY-3").Position = 1
    End With
    
    pt.PivotFields("EMPLOYEE").ShowDetail = False
    
    'hide blanks
    For Each pi In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (pi.Value = "(blank)") Then
            pi.Visible = False
        Else: pi.Visible = True
        End If
    Next pi
    
    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    pt.PivotSelect ("")
    Charts.Add
    
    If chartatlastcol Then
        Set cLoc = ws.Cells(1, LastCol(ws) + 1)
    End If
    
    ActiveChart.Location Where:=xlLocationAsObject, name:=pt.Parent.name
    ActiveChart.Parent.Left = cLoc.Left
    ActiveChart.Parent.Top = cLoc.Top
    ActiveChart.ApplyLayout (3)
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.ShowAllFieldButtons = False
    ActiveChart.HasTitle = True
    ActiveChart.chartTitle.Text = cTitle
    ActiveChart.Axes(xlValue).MaximumScale = 0.2
   
    Selection.Format.TextFrame2.TextRange.Characters.Text = cTitle

    'chart format goes here
    Dim ci As Series
    Dim num As Double
    Dim s As String

    For Each ci In ActiveChart.SeriesCollection
        s = ci.name
        num = CDbl(Mid(ci.name, 3, Len(s)))
        
        If num > 0 Then
            With ci.Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorAccent2
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.5
                .Transparency = 1 - (Abs(num) / 3)
                .Solid
            End With
            
        ElseIf num < 0 Then
            With ci.Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorAccent1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.5
                .Transparency = 1 - (Abs(num) / 3)
                .Solid
            End With
        ElseIf num = 0 Then
            With ci.Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorText1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = Abs(num)
                .Transparency = 0.5
                .Solid
            End With
        End If
    Next ci
        
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    'ActiveChart.Axes(xlValue).MaximumScale = 1
    
    pasteCol = LastCol(ws) + 2
    
    With [A1]
        pt.TableRange2.Copy
        ws.Cells(1, pasteCol).PasteSpecial xlPasteValues
    End With

    pasteRow = 1
    pasteCol = LastCol(ws) + 2
    
    ws.Cells(pasteRow, pasteCol).Value = "CY0 Rates: "
    pasteRow = pasteRow + 1
    ws.Cells(pasteRow, pasteCol).Value = "Cytotech CY0 Rate: "
    ws.Cells(pasteRow, pasteCol + 1).FormulaR1C1 = "=SUM(IFERROR(GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Cytotechnologist"", ""QUALITY CODE"", ""CY0""),0),IFERROR(GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Technologist"", ""QUALITY CODE"", ""CY0""),0))/SUM(IFERROR(GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Cytotechnologist""),0),IFERROR(GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Technologist""),0))"
    ws.Cells(pasteRow + 1, pasteCol).Value = "Resident CY0 Rate: "
    ws.Cells(pasteRow + 1, pasteCol + 1).FormulaR1C1 = "=GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Resident"", ""QUALITY CODE"", ""CY0"")"
    ws.Cells(pasteRow + 2, pasteCol).Value = "Fellow CY0 Rate: "
    ws.Cells(pasteRow + 2, pasteCol + 1).FormulaR1C1 = "=GetPivotData(""CASE NUMBER"", R3C1, ""EMPLOYEE TYPE"", ""Fellow"", ""QUALITY CODE"", ""CY0"")"
    
    For i = pasteRow To pasteRow + 2
        ws.Cells(i, pasteCol + 2).FormulaR1C1 = "=RC[-2]&RC[-1]"
        ws.Cells(i, pasteCol + 3).Value = ws.Cells(i, pasteCol + 2).Value
        ws.Cells(i, pasteCol + 4).FormulaR1C1 = "=IF(ISERROR(RC[-3]),"""",RC[-1])"
        
        If i = pasteRow Then
            With ActiveSheet.PivotTables("PTCTPathAgreement").PivotFields("Count of CASE NUMBER")
                .Calculation = xlPercentOfRow
                .NumberFormat = "0.00%"
            End With
        End If
    Next i
   
    ws.ChartObjects(1).Activate
    ActiveChart.Shapes.AddLabel(msoTextOrientationHorizontal, 12, 5, 128, 12).Select
    Selection.Formula = "='CTPathAgreement'!R2C" & LastCol(ws)
    ActiveChart.Shapes.AddLabel(msoTextOrientationHorizontal, 12, 10, 128, 12).Select
    Selection.Formula = "='CTPathAgreement'!R3C" & LastCol(ws)
    ActiveChart.Shapes.AddLabel(msoTextOrientationHorizontal, 12, 15, 128, 12).Select
    Selection.Formula = "='CTPathAgreement'!R4C" & LastCol(ws)
   
    With ActiveSheet.PivotTables("PTCTPathAgreement").PivotFields("Count of CASE NUMBER")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0.00%"
    End With
    
    With ActiveChart.Parent
        .Height = 600 ' resize 2.5 pt at 72 ppi.
        .Width = 800 ' resize 4.0 pt at 72 ppi.
    End With
    
    'add slicer buttons
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PTCTPathAgreement"), _
        "EMPLOYEE").Slicers.Add ActiveSheet, , "EMPLOYEE 3", "EMPLOYEE", 197.25, 495.75 _
        , 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PTCTPathAgreement"), _
        "EMPLOYEE TYPE").Slicers.Add ActiveSheet, , "EMPLOYEE TYPE 2", "EMPLOYEE TYPE" _
        , 234.75, 533.25, 144, 187.5
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("PTCTPathAgreement"), _
        "DIAGNOSIS CATEGORY").Slicers.Add ActiveSheet, , "DIAGNOSIS CATEGORY 2", _
        "DIAGNOSIS CATEGORY", 272.25, 570.75, 144, 187.5

    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    Range("N50").Value = "Change chart scale by right clicking on the Y-axis > Format Axis > Bounds, Maximum."

End Sub


Private Sub MultiSheetSub()
    'automatically selected by CleanUp() when multiple data sheets are available
    UnmergeAll
    DeleteEmptyRows
    CopyRangeFromMultiWorksheets
    HideSheets
    UpdateHPVResults
    DeleteHPVLines
    DeleteDuplicateInterpretations
    'InsertHPVOverall               'commented to allow run from updatehpvresults based on mankato
    SortData
    GeneratePT
    RowSizeZoom
End Sub

Private Sub SingleSheetSub()
    'automatically selected by CleanUp() when only one data sheet is available
    UnmergeAll
    DeleteEmptyRows
    RenameSheet
    UpdateHPVResults
    DeleteHPVLines
    DeleteDuplicateInterpretations
    'InsertHPVOverall               'commented to allow run from updatehpvresults based on mankato
    SortData
    GeneratePT
    RowSizeZoom
End Sub

Private Sub QuickRecopy()
    'for use when data already is unmerged and empty rows deleted
    CopyRangeFromMultiWorksheets
    HideSheets
    UpdateHPVResults
    DeleteHPVLines
    DeleteDuplicateInterpretations
    'InsertHPVOverall               'commented to allow run from updatehpvresults based on mankato
    SortData
    RowSizeZoom
End Sub

Sub GeneratePT()
    'generates pivot tables
    PTInterpTotals
    PTBenchmarks
    PTHPVbyDx
    PTASCUSHPV
    PTCTAgreement
    PTCTPathAgreement
End Sub

