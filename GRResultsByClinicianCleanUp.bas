Attribute VB_Name = "GRResultsByClinicianCleanUp"
Option Explicit

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

Sub DeleteTitles()
    ' deletes report titles from first and last row so they don't interfere with data
    Dim ws As Worksheet, lr As Long

    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    lr = LastRow(ws)
    
    With ws
        If Range("A1").Value = "PathDx Cytology Results by Clinician" Then
            Range("A1").SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
        If (Cells(LastRow(ws), 1).Value = "Report Title: PathDX Cytology Results by Clinician") Then
            Cells(LastRow(ws), 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End With
    
End Sub

Sub ValidateRequestingDoctor()
    Dim ws As Worksheet, i As Long
    Dim reqDocCol As Long
    
    Set ws = ActiveWorkbook.Worksheets("Sheet1")
    
    With Application
        '.Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    'find REQUESTING DOCTOR column
    For i = 1 To LastCol(ws)
        If (ws.Cells(1, i).Value = "REQUESTING DOCTOR") Then
            reqDocCol = i
            Exit For
        ElseIf (i = LastCol(ws)) And (reqDocCol = 0) Then
            MsgBox "Could not find REQUESTING DOCTOR field.  Verify that there is a column named REQUESTING DOCTOR in Sheet1."
            Exit Sub
        End If
    Next i
    
    'find REQUESTING DOCTOR VALIDATED column, if it exists, clear it
    For i = 1 To LastCol(ws)
        If (ws.Cells(1, i).Value = "REQUESTING DOCTOR VALIDATED") Or (ws.Cells(1, i).Value = "REQ DOC LNAME") Then
            ws.Columns(i).Clear
        End If
    Next i
    
    'append space onto end of REQUESTING DOCTOR value if the last char is not a space
    For i = 2 To LastRow(ws)
        If (Right(ws.Cells(i, reqDocCol).Value, 1) <> " ") Then
            ws.Cells(i, reqDocCol).Value = ws.Cells(i, reqDocCol).Value & " "
        End If
    Next i
        
    'set up requesting doctor column for data validation
    Dim relRef As Long
    
    ws.Cells(1, LastCol(ws) + 1).Value = "REQUESTING DOCTOR VALIDATED"
    relRef = LastCol(ws) - reqDocCol
    ws.Cells(2, LastCol(ws)).Formula = "=LEFT(RC[-" & relRef & "], FIND(CHAR(44), RC[-" & relRef & "]))" _
        & "&MID(RC[-" & relRef & "], FIND(CHAR(44),RC[-" & relRef & "])+1,FIND(CHAR(32),RC[-" & relRef & "],FIND(CHAR(44),RC[-" & relRef & "])+2)-FIND(CHAR(44),RC[-" & relRef & "]))"
    
    ws.Cells(1, LastCol(ws) + 1).Value = "REQ DOC LNAME"
    relRef = LastCol(ws) - reqDocCol
    ws.Cells(2, LastCol(ws)).Formula = "=LEFT(RC[-" & relRef & "], FIND(CHAR(44), RC[-" & relRef & "])-1)"
    
    'find first comma to end last name, find first space to start first name, position of second space minus position of first space +1 is length of first name
    '=LEFT(RC[-7], FIND(CHAR(44), RC[-7]))&MID(RC[-7], FIND(CHAR(32),RC[-7]),FIND(CHAR(32),RC[-7],FIND(CHAR(32),RC[-7])+1)-FIND(CHAR(32),RC[-7]))
    
    'find first comma to end last name, find first comma to start first name, position of last space minus position of comma +2 is length of first name
    '=LEFT(RC[-7], FIND(CHAR(44), RC[-7]))
    '      &MID(RC[-7],              FIND(CHAR(44),RC[-7             ])+1,FIND(CHAR(32),RC[-7             ],FIND(CHAR(44),RC[-7             ])+2)-FIND(CHAR(32),RC[-7]))
    
    'fill to last row
    ws.Cells(2, LastCol(ws) - 1).AutoFill Destination:=Range(ws.Cells(2, LastCol(ws) - 1), ws.Cells(LastRow(ws), LastCol(ws) - 1))
    ws.Cells(2, LastCol(ws)).AutoFill Destination:=Range(ws.Cells(2, LastCol(ws)), ws.Cells(LastRow(ws), LastCol(ws)))

    
    'copy and paste values, comment out to troubleshoot formula
    With ws.Range(ws.Cells(2, LastCol(ws) - 1), ws.Cells(LastRow(ws), LastCol(ws)))
        .Copy
        .PasteSpecial xlPasteValues
    End With
    
    With Application
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
       
End Sub

Sub MayoDocFilters()

    Dim enterSub As Integer
    enterSub = MsgBox("Would you like to filter Mayo data based on a last name list?", vbYesNo)
    If enterSub = vbNo Then Exit Sub
    
    Dim wb As Workbook, ws As Worksheet, ptws As Worksheet, pt As PivotTable, pi As PivotItem, pf As PivotField, i As Long
    Dim reqDocCol As Long, fSheetName As String, fws As Worksheet
            
    fSheetName = "MayoDocFilters"
        
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("Sheet1")
    Set ptws = wb.Worksheets("MayoClinBenchmarks")

    With Application
        '.Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
        
    'find REQUESTING DOCTOR VALIDATED column
    For i = 1 To LastCol(ws)
        If ws.Cells(1, i).Value = "REQUESTING DOCTOR VALIDATED" Then
            reqDocCol = i
            Exit For
        ElseIf (i = LastCol(ws)) And (reqDocCol = 0) Then
            MsgBox "Could not find REQUESTING DOCTOR VALIDATED field.  Please re-run BuildPT or ValidateRequestingDoctor."
            Exit Sub
        End If
    Next i
    
    Dim n As Long, count As Long, qF As Boolean
    n = wb.Worksheets.count
    count = 1
    Set ws = Nothing
        
    For Each ws In wb.Sheets
        If ws.name = fSheetName Then
            Set fws = ws
            Exit For
        ElseIf (count = n) And Not qF Then
            wb.Sheets.Add
            ActiveSheet.name = fSheetName
            ws.Range("A1").Value = "REQUESTING DOCTOR VALIDATED"
            
            MsgBox "MayoDocFilters sheet created.  Please paste Mayo doctor last names in the column indicated and re-run BuildPT or PTMayoClinicianResults."
            Exit Sub
            
        End If
        count = count + 1
    Next ws
    
    'other ideas for string processing: lcase (works without), remove spaces and hyphens
    For Each pt In ptws.PivotTables
        With pt.PivotFields("REQ DOC LNAME")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = WorksheetFunction.CountIf(Range(fws.Cells(2, 1), fws.Cells(LastRow(fws), 1)), pi.name) > 0
            Next pi
        End With
    Next pt
    
    'way too slow
    'For Each pt In ptws.PivotTables
    '    For Each pi In pt.PivotFields("REQUESTING DOCTOR VALIDATED").PivotItems
    '        For i = 2 To LastRow(wb.Worksheets(fSheetName))
    '            If Left(pi.Value, InStr(pi.Value, Chr(44))) = Cells(i, 1) Then
    '                pi.Visible = True
    '            Else
    '                pi.Visible = False
    '            End If
    '        Next i
    '    Next pi
    'Next pt
    
    With Application
        '.Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
        
End Sub

Sub PTMayoClinicianResults()
'
' PTBenchmarks Macro
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, pt As PivotTable, pi As PivotItem, pt2 As PivotTable, Sheet1 As Worksheet
    Dim ws As Worksheet, name As String, lcol As Long, lrow As Long, i As Integer, j As Integer
    Dim agusCount As Long
            
    name = "MayoClinBenchmarks"         'worksheet name, used to construct pivot table names
    
    Set Sheet1 = Worksheets("Sheet1")
    
    DeleteTitles
    ValidateRequestingDoctor
    
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(name).Delete
    Application.DisplayAlerts = True
    Sheet1.Activate
    
    On Error GoTo 0
    
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Sheet1.Range("A2").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(79, 129, 189)
    End With
    
    Set ws = Sheets(name)
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:="PT" & name & "Count")
   
    With pt.PivotFields("HOSPITAL CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("REQUESTING DOCTOR VALIDATED")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("COLLECTION DATE")
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
    With pt.PivotFields("NORMAL / ABNORMAL")
        .Orientation = xlColumnField
        .Position = 1
    End With
    pt.PivotFields("REQUESTING DOCTOR VALIDATED").ShowDetail = False
    
    With pt.PivotFields("REQ DOC LNAME")
        .Orientation = xlPageField
        .Position = 1
    End With

    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    lcol = LastCol(ws) + 2
    Set pt2 = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=ws.Cells(1, lcol), TableName:="PTBenchmarksPercent")
    
    With pt2.PivotFields("HOSPITAL CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt2.PivotFields("REQUESTING DOCTOR VALIDATED")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt2.PivotFields("COLLECTION DATE")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt2.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 4
    End With
    pt2.PivotFields("REQUESTING DOCTOR VALIDATED").ShowDetail = False
    
    'Diagnosis Category Setup
    With pt2.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt2.PivotFields("NORMAL / ABNORMAL")
        .Orientation = xlColumnField
        .Position = 1
    End With

    With pt2.PivotFields("REQ DOC LNAME")
        .Orientation = xlPageField
        .Position = 1
    End With

    'add interp count by case number and collapse to employee
    pt2.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    'change second PT to percentage
    With pt2.PivotFields("Count of CASE NUMBER")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0.00%"
    End With

    'count agus items, and group them if more than one
    On Error Resume Next
    Application.PivotTableSelection = True
    Dim pTable As PivotTable, dxCat As PivotField, hCode As PivotField
    Dim itm As PivotItem '   < added
    Dim unionrng As Range
    
    For Each itm In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (itm.name = "GYN NIL" Or itm.name = "GYNNOEC" Or itm.name = "GYN ORG" Or itm.name = "GYN REAC") Then
            If unionrng Is Nothing Then
                Set unionrng = itm.LabelRange
            Else
                Set unionrng = Union(unionrng, itm.LabelRange)
            End If
        End If
    Next itm
    unionrng.Group
    
    'reset unionrng
    Set unionrng = Nothing
    
    'loop through agus to group them
    For Each itm In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (itm.name = "GYN AGUS" Or itm.name = "GYN AIS") Then
            If unionrng Is Nothing Then
                Set unionrng = itm.LabelRange
            Else
                Set unionrng = Union(unionrng, itm.LabelRange)
            End If
        End If
    Next itm
    unionrng.Group
    
    'loop through each pivotitem in each ptable and rename nil and agus categories, and hide ngyn categories
    For Each pTable In Sheets(name).PivotTables()
        Set dxCat = pTable.PivotFields("DIAGNOSIS CATEGORY2")
        Set hCode = pTable.PivotFields("HOSPITAL CODE")
        For Each itm In dxCat.PivotItems
            If (itm.name = "Group1") Then
                itm.name = "NIL"
            ElseIf (itm.name = "Group2") Then
                itm.name = "AGUS"
            ElseIf (Left(itm.name, 4) = "NGYN") Then
                itm.Visible = False
            End If
        Next itm
        
        'hide mml cases for mayo table
        For Each itm In hCode.PivotItems
            If (itm.name = "2MML") Then
                itm.Visible = False
            End If
        Next itm
        
        With pTable.PivotFields("DIAGNOSIS CATEGORY2")
            On Error Resume Next
            .ShowDetail = False
            .PivotItems("GYN CANCER").Position = 1
            .PivotItems("GYN AGUS").Position = 1
            .PivotItems("GYN AIS").Position = 1
            .PivotItems("AGUS").Position = 1
            .PivotItems("GYN HSIL").Position = 1
            .PivotItems("GYN LSIL").Position = 1
            .PivotItems("GYN ASCH").Position = 1
            .PivotItems("GYN ASCUS").Position = 1
            .PivotItems("NIL").Position = 1
            .PivotItems("GYN UNSAT").Position = 1
        End With
        
        With pTable.PivotFields("NORMAL / ABNORMAL")
            '.ShowDetail = True
            .PivotItems("ABNORMAL").Position = 1
            .PivotItems("NORMAL").Position = 1
            .Subtotals(1) = False
        End With
        
        'hide blanks
        pTable.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
        
    Next pTable
    
'    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
    MayoDocFilters
    
End Sub

Sub PTMMLClinicianResults()
'
' PTBenchmarks Macro
'

' https://www.mrexcel.com/forum/excel-questions/785527-macro-create-pivot-table-dynamic-data-range.html

    Dim PCache As PivotCache, pt As PivotTable, pi As PivotItem, pt2 As PivotTable, Sheet1 As Worksheet
    Dim ws As Worksheet, name As String, lcol As Long, lrow As Long, i As Integer, j As Integer
    Dim agusCount As Long
            
    name = "MMLClinBenchmarks"         'worksheet name, used to construct pivot table names
    
    Set Sheet1 = Worksheets("Sheet1")
    
    DeleteTitles
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(name).Delete
    Application.DisplayAlerts = True
    Sheet1.Activate
    
    On Error GoTo 0
    
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=1, SourceData:=Sheet1.Range("A2").CurrentRegion.Address)
    Worksheets.Add
    With ActiveSheet
        .name = name
        .Tab.Color = RGB(128, 100, 162)
    End With
    
    Set ws = Sheets(name)
    Set pt = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=Range("A1"), TableName:="PT" & name & "Count")
   
    With pt.PivotFields("HOSPITAL CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt.PivotFields("WARD NAME")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt.PivotFields("REQUESTING DOCTOR VALIDATED")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt.PivotFields("COLLECTION DATE")
        .Orientation = xlRowField
        .Position = 4
    End With
    With pt.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    
    'Diagnosis Category Setup
    With pt.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt.PivotFields("NORMAL / ABNORMAL")
        .Orientation = xlColumnField
        .Position = 1
    End With
    pt.PivotFields("REQUESTING DOCTOR VALIDATED").ShowDetail = False

    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    lcol = LastCol(ws) + 2
    Set pt2 = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=ws.Cells(1, lcol), TableName:="PTBenchmarksPercent")
    
    With pt2.PivotFields("HOSPITAL CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt2.PivotFields("WARD NAME")
        .Orientation = xlRowField
        .Position = 2
    End With
    With pt2.PivotFields("REQUESTING DOCTOR VALIDATED")
        .Orientation = xlRowField
        .Position = 3
    End With
    With pt2.PivotFields("COLLECTION DATE")
        .Orientation = xlRowField
        .Position = 4
    End With
    With pt2.PivotFields("CASE NUMBER")
        .Orientation = xlRowField
        .Position = 5
    End With
    pt2.PivotFields("REQUESTING DOCTOR VALIDATED").ShowDetail = False
    
    'Diagnosis Category Setup
    With pt2.PivotFields("DIAGNOSIS CATEGORY")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pt2.PivotFields("NORMAL / ABNORMAL")
        .Orientation = xlColumnField
        .Position = 1
    End With

    'add interp count by case number and collapse to employee
    pt2.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    'change second PT to percentage
    With pt2.PivotFields("Count of CASE NUMBER")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0.00%"
    End With

    'count agus items, and group them if more than one
    On Error Resume Next
    Application.PivotTableSelection = True
    Dim pTable As PivotTable, dxCat As PivotField, hCode As PivotField
    Dim itm As PivotItem '   < added
    Dim unionrng As Range
    
    For Each itm In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (itm.name = "GYN NIL" Or itm.name = "GYNNOEC" Or itm.name = "GYN ORG" Or itm.name = "GYN REAC") Then
            If unionrng Is Nothing Then
                Set unionrng = itm.LabelRange
            Else
                Set unionrng = Union(unionrng, itm.LabelRange)
            End If
        End If
    Next itm
    unionrng.Group
    
    'reset unionrng
    Set unionrng = Nothing
    
    'loop through agus to group them
    For Each itm In pt.PivotFields("DIAGNOSIS CATEGORY").PivotItems
        If (itm.name = "GYN AGUS" Or itm.name = "GYN AIS") Then
            If unionrng Is Nothing Then
                Set unionrng = itm.LabelRange
            Else
                Set unionrng = Union(unionrng, itm.LabelRange)
            End If
        End If
    Next itm
    unionrng.Group
    
    'loop through each pivotitem in each ptable and rename nil and agus categories, and hide ngyn categories
    For Each pTable In Sheets(name).PivotTables()
        Set dxCat = pTable.PivotFields("DIAGNOSIS CATEGORY2")
        Set hCode = pTable.PivotFields("HOSPITAL CODE")
        For Each itm In dxCat.PivotItems
            If (itm.name = "Group1") Then
                itm.name = "NIL"
            ElseIf (itm.name = "Group2") Then
                itm.name = "AGUS"
            ElseIf (Left(itm.name, 4) = "NGYN") Then
                itm.Visible = False
            End If
        Next itm
        
        'hide mayo cases for mml table
        For Each itm In hCode.PivotItems
            If (itm.name <> "2MML") Then
                itm.Visible = False
            End If
        Next itm
        
        With pTable.PivotFields("DIAGNOSIS CATEGORY2")
            On Error Resume Next
            .ShowDetail = False
            .PivotItems("GYN UNSAT").Position = 1
            .PivotItems("GYN CANCER").Position = 1
            .PivotItems("GYN AGUS").Position = 1
            .PivotItems("GYN AIS").Position = 1
            .PivotItems("AGUS").Position = 1
            .PivotItems("GYN HSIL").Position = 1
            .PivotItems("GYN LSIL").Position = 1
            .PivotItems("GYN ASCH").Position = 1
            .PivotItems("GYN ASCUS").Position = 1
            .PivotItems("NIL").Position = 1
        End With
        
        With pTable.PivotFields("NORMAL / ABNORMAL")
            '.ShowDetail = True
            .PivotItems("ABNORMAL").Position = 1
            .PivotItems("NORMAL").Position = 1
            .Subtotals(1) = False
        End With
        
        'hide blanks
        pTable.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
        pTable.PivotFields("DIAGNOSIS CATEGORY").PivotItems("(blank)").Visible = False
    Next pTable
    
    ws.Rows(1).SpecialCells(xlCellTypeBlanks).Select
    
    'zASCtoSIL deleted to next line
    
'    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
     
End Sub

Sub RowSizeZoom()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    ws.Range("A2:A" & ws.Rows.count).rowHeight = 12.75
    ws.Activate
    ActiveWindow.Zoom = 85
  Next ws
  
End Sub

Sub BuildPT()
    PTMayoClinicianResults
    PTMMLClinicianResults
    RowSizeZoom
End Sub

