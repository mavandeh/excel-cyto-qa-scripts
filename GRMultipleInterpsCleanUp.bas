Attribute VB_Name = "GRMultipleInterpsCleanUp"
' This file is meant for cleaning up PathDx Cytology Multiple Interpretations Report (GYN Version)
' Author: Mark A. VandeHaar, SCT(ASCP)
' See

Option Explicit

Sub UnmergeAll()
  Dim currentSheet As Worksheet
  For Each currentSheet In Worksheets
    currentSheet.Cells.unmerge
  Next
End Sub

Sub DeleteEmptyRows()
  Dim currentSheet As Worksheet
  For Each currentSheet In Worksheets
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

Sub CopyRangeFromMultiWorksheets()

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
    DestSh.Name = "Data"

    ' Loop through all worksheets and copy the data to the summary worksheet.
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Name <> DestSh.Name Then

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

    Application.Goto DestSh.Cells(1)

    ' AutoFit the column width in the summary sheet.
    DestSh.Columns.AutoFit

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Sub SortData()
    ' created from macro recording and modified
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("A1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("B1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "HPV,TPRPS,TPRPD,STHPV,DTHPV,STPCO,DTPCO", DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("P1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A:Y")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    
End Sub

Sub UpdateHPVResults()
  
  With Application
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .ReferenceStyle = xlR1C1
  End With

    Dim col As Integer
    For col = 20 To 22
    
      Dim cell As Range
      On Error Resume Next
      For Each cell In Columns(col).SpecialCells(xlCellTypeBlanks).Areas
         cell.FormulaArray = "=INDEX(C,MATCH(RC1&""HPV"",C1&C2,0))"
         cell.Value = cell.Value ' comment this line for formula troubleshooting
      Next cell
    Next col
    
  With Application
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
  End With
    
End Sub

Sub RenameSheet()
    ActiveSheet.Name = "Data"
End Sub

Sub DeleteHPVLines()

    ' Modified from https://danwagner.co/how-to-delete-rows-with-range-autofilter/
    Dim wksData As Worksheet
    Dim lngLastRow As Long
    Dim rngData As Range

    Set wksData = ThisWorkbook.Worksheets("Data")
    
    With wksData
        lngLastRow = .Range("A" & .Rows.Count).End(xlUp).Row
        Set rngData = .Range("A1:X" & lngLastRow)
    End With
    
    Application.DisplayAlerts = False
        With rngData
            ' Filter for HPV in the test column (#2)
            .AutoFilter Field:=2, Criteria1:="HPV"
            ' Delete visible rows, keep header
            .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With
    Application.DisplayAlerts = True
    
    'Turn off the AutoFilter
    With wksData
        .AutoFilterMode = False
        If .FilterMode = True Then
            .ShowAllData
        End If
    End With
  
End Sub

Sub DeleteDuplicateInterpretations()

  ' Original code utilizing LastRow() Function from MSDN above.

  Dim ws As Worksheet
  Dim lr As Long
    
  Set ws = Worksheets("Data")
  lr = LastRow(ws)
  
  'create case-employee column (Y=A&I)
  Range("Y1").Value = "CASE_EMPLOYEE"
  Range("Y2").Formula = "=A2&I2"
  Range("Y2").AutoFill Destination:=Range("Y2:Y" & lr)
  
  'sort by case-person ascending (Y) and then by interp date descending (P)
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("Y1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("P1") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A:Y")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
  'remove duplicates based on case-person column (will delete second entries leaving the most recent which is sorted at top)
  Range("A1:Y" & lr).RemoveDuplicates Columns:=Array(25)
  
End Sub

Sub RowSizeZoom()
  Dim ws As Worksheet
  For Each ws In Worksheets
    ws.Range("A2:A" & ws.Rows.Count).RowHeight = 12.75
  Next ws
  ActiveWindow.Zoom = 70
End Sub

Sub MultipleInterpsCleanUpMerged()
  
  UnmergeAll
  DeleteEmptyRows
  CopyRangeFromMultiWorksheets
  UpdateHPVResults
  DeleteHPVLines
  DeleteDuplicateInterpretations
  SortData
  RowSizeZoom
  
End Sub

Sub MultipleInterpsCleanUpNotMergedOneSheet()
  
  RenameSheet
  UpdateHPVResults
  DeleteHPVLines
  DeleteDuplicateInterpretations
  SortData
  RowSizeZoom
End Sub
