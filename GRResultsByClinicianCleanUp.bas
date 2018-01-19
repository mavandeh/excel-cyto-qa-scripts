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
    With pt.PivotFields("REQUESTING DOCTOR")
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
    pt.PivotFields("REQUESTING DOCTOR").ShowDetail = False

    'add interp count by case number and collapse to employee
    pt.AddDataField pt.PivotFields("CASE NUMBER"), "Count of CASE NUMBER", xlCount
    
    lcol = LastCol(ws) + 2
    Set pt2 = ActiveSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=ws.Cells(1, lcol), TableName:="PTBenchmarksPercent")
    
    With pt2.PivotFields("HOSPITAL CODE")
        .Orientation = xlRowField
        .Position = 1
    End With
    With pt2.PivotFields("REQUESTING DOCTOR")
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
    pt2.PivotFields("REQUESTING DOCTOR").ShowDetail = False
    
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

    'ws.Rows(1).SpecialCells(xlCellTypeBlanks, XlCellType).Select
    
    'zASCtoSIL deleted to next line
    
'    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    
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
    With pt.PivotFields("REQUESTING DOCTOR")
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
    pt.PivotFields("REQUESTING DOCTOR").ShowDetail = False

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
    With pt2.PivotFields("REQUESTING DOCTOR")
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
    pt2.PivotFields("REQUESTING DOCTOR").ShowDetail = False
    
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
        
    Next pTable
    
    ws.Rows(1).SpecialCells(xlCellTypeBlanks).Select
    
    'zASCtoSIL deleted to next line
    
'    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
     
End Sub

Sub RowSizeZoom()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    ws.Range("A2:A" & ws.Rows.Count).rowHeight = 12.75
    ws.Activate
    ActiveWindow.Zoom = 85
  Next ws
  
End Sub

Sub BuildPT()
    PTMayoClinicianResults
    PTMMLClinicianResults
    RowSizeZoom
End Sub

