Attribute VB_Name = "AggregateCytoHisto"
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

Function lastCol(sh As Worksheet)
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

Sub AggregateCytoHisto()

    Dim wbSource As Workbook
    Dim shTarget As Worksheet
    Dim shSource As Worksheet
    Dim shValidation As Worksheet
    Dim strFilePath As String
    Dim strPath As String

    ' Initialize some variables and
    ' get the folder path that has the files
    Set shTarget = ThisWorkbook.Sheets("Sheet1")
    strPath = GetPath

    ' Make sure a folder was picked.
    If Not strPath = vbNullString Then

        ' Get all the files from the folder
        strFilePath = Dir$(strPath & "*.xls*", vbNormal)

        Do While Not strFilePath = vbNullString

            ' Open the file and get the source sheet
            Set wbSource = Workbooks.Open(strPath & strFilePath)
            Set shSource = wbSource.Sheets("Sheet1")

            'Copy the data
            Call CopyData(shSource, shTarget)

            'Close the workbook and move to the next file.
            wbSource.Close False
            strFilePath = Dir$()
        Loop
    End If
    
    DeleteHeaders
    RowSizeZoom

End Sub

Private Sub CopyData(ByRef shSource As Worksheet, shTarget As Worksheet)

    Dim lRowSource As Long, lRowTarget As Long, lColSource As Long, lColTarget As Long
    
    'Determine the last row.
    'lRow = shTarget.Cells(shTarget.Rows.Count, 3).End(xlUp).Row + 1
    lRowSource = LastRow(shSource)
    lRowTarget = LastRow(shTarget)
       
    'Determine the last column.
    lColSource = lastCol(shSource)
    'lColTarget = LastCol(shTarget)

    'Copy range from source to target
    shSource.UsedRange.Copy
    'shSource.Range("A1", shSource.Cells(lRowSource & lColSource)).Select
    'Selection.Copy
    shTarget.Cells(lRowTarget + 1, 1).PasteSpecial xlPasteValuesAndNumberFormats

    ' Reset the clipboard.
    Application.CutCopyMode = xlCopy

End Sub


' Fucntion to get the folder path
Function GetPath() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .ButtonName = "Select a folder"
        .Title = "Folder Picker"
        .AllowMultiSelect = False

        'Get the folder if the user does not hot cancel
        If .Show Then GetPath = .SelectedItems(1) & "\"

    End With

End Function

Private Sub DeleteHeaders()
    Dim lrow As Long
    Dim ws As Worksheet
    Dim strSearch As String
    
    
    
    '~~> Set this to the relevant worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    '~~> Search Text
    strSearch = "PathDx"
    strSearch2 = "Report"
    strSearch3 = "Case Number"

    With ws
        '~~> Remove any filters
        .AutoFilterMode = False

        lrow = .Range("A" & .Rows.Count).End(xlUp).Row

        With .Range("A1:A" & lrow)
            .AutoFilter Field:=1, Criteria1:="=*" & strSearch & "*"
            .AutoFilter Field:=1, Criteria1:="=*" & strSearch2 & "*"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
    End With
    
    With ws
        '~~> Remove any filters
        .AutoFilterMode = False

        lrow = .Range("A" & .Rows.Count).End(xlUp).Row

        With .Range("A2:A" & lrow)
            .AutoFilter Field:=1, Criteria1:="=*" & strSearch3 & "*"
            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End With

        '~~> Remove any filters
        .AutoFilterMode = False
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
