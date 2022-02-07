Attribute VB_Name = "AggregateDAF"
Option Explicit

Sub AggregateDAF()

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
        strFilePath = Dir$(strPath & "*.xlsx", vbNormal)

        Do While Not strFilePath = vbNullString

            ' Open the file and get the source sheet
            Set wbSource = Workbooks.Open(strPath & strFilePath)
            Set shSource = wbSource.Sheets("015201")
            Set shValidation = wbSource.Sheets("Rev-Sig")

            'Copy the data
            Call CopyData(shSource, shTarget, shValidation)

            'Close the workbook and move to the next file.
            wbSource.Close False
            strFilePath = Dir$()
        Loop
    End If

End Sub

Function getEnd(rng As Range)
    

End Function


' Procedure to copy the data.
Sub CopyData(ByRef shSource As Worksheet, shTarget As Worksheet, shValidation As Worksheet)

    Dim nameRng As String, yearRng As String, monthRng As String
    Dim name As String, nameOut As String, nameSpaceLast As Long, nameSpaceFirst As Long
    Dim GynSldRng As String, GynHrsRng As String
    Dim ngcSldRng As String, ngcHrsRng As String, ngcHrsSum As Long
    Dim i As Long, iEnd As Long, lRow As Long
        
    'find the most recent version entry, and then get the ver number
    iEnd = 10               'how many cells to look through for versions (start with last)
    For i = 0 To iEnd
            
        If shValidation.Range("C" & iEnd - i).Value = "037" Then
            nameRng = "B7"
            GynSldRng = "AG11"
            GynHrsRng = "AG14"
            ngcSldRng = "AG12"
            ngcHrsRng = "AG15:AG17"
            yearRng = "AE7"
            monthRng = "W7"
            ngcHrsSum = shSource.Range("AG15").Value + shSource.Range("AG16").Value + shSource.Range("AG17").Value
            Exit For
            
        ElseIf shValidation.Range("C" & iEnd - i).Value = "035" Then
            nameRng = "B8"
            GynSldRng = "AG12"
            GynHrsRng = "AG15"
            ngcSldRng = "AG13"
            ngcHrsRng = "AG16:AG18"
            yearRng = "AE8"
            monthRng = "W8"
            ngcHrsSum = shSource.Range("AG16").Value + shSource.Range("AG17").Value + shSource.Range("AG18").Value
            Exit For
            
        End If
    Next i
    
    'set up first row
    If Not IsEmpty(Range("A1")) Then
        shTarget.Cells(1, 1).Value = "Month"
        shTarget.Cells(1, 2).Value = "Quarter"
        shTarget.Cells(1, 3).Value = "Name"
        shTarget.Cells(1, 4).Value = "Total GYN Slides"
        shTarget.Cells(1, 5).Value = "GYN Hours"
        shTarget.Cells(1, 6).Value = "Primary GYN Slides"
        shTarget.Cells(1, 7).Value = "GYN Slides per Hour"
        shTarget.Cells(1, 8).Value = "Total Non-GYN Slides"
        shTarget.Cells(1, 9).Value = "Non-GYN Hours"
        shTarget.Cells(1, 10).Value = "Non-GYN Slides per Hour"
        shTarget.Cells(1, 11).Value = "Tech Number"
        shTarget.Cells(1, 12).Value = "Tech Initials"
        shTarget.Cells(1, 13).Value = "Total Slides per Hour"
    End If
    
    'Determine the last row.
    lRow = shTarget.Cells(shTarget.Rows.Count, 3).End(xlUp).Row + 1
        
    'Copy Name:
    'If there isn't a comma in the name indicating Lname, Fname format, then
    'trim value of the name range, then extract last name, first name, and concatenate them
    'and place the name in proper column
    If InStr(shSource.Range(nameRng).Value, ",") Then
        nameOut = Trim(shSource.Range(nameRng).Value)
    Else
        name = Trim(shSource.Range(nameRng).Value)
        nameSpaceLast = InStrRev(name, " ")
        nameSpaceFirst = InStr(name, " ")
        nameOut = Right(name, Len(name) - nameSpaceLast) & ", " & Left(name, nameSpaceFirst - 1)
    End If
    shTarget.Cells(lRow, 3).Value = nameOut

    'Copy month and year.  Typos lead to validation errors.
    'User must correct them, save the file, delete the data rows generated from that folder and re-run.
    shTarget.Cells(lRow, 1).Value = shSource.Range(monthRng).Value & " " & shSource.Range(yearRng).Value
    shTarget.Cells(lRow, 2).Value = DatePart("q", shTarget.Cells(lRow, 1).Value) & "Q-" & shSource.Range(yearRng)

    'copy primary gyn slides and hours
    shSource.Range(GynSldRng).Copy
    shTarget.Cells(lRow, 4).PasteSpecial xlPasteValuesAndNumberFormats
    shSource.Range(GynHrsRng).Copy
    shTarget.Cells(lRow, 5).PasteSpecial xlPasteValuesAndNumberFormats
    shTarget.Cells(lRow, 7).FormulaR1C1 = "=RC[-3]/RC[-2]"
    
    'Copy ngc slides and hours
    shSource.Range(ngcSldRng).Copy
    shTarget.Cells(lRow, 8).PasteSpecial xlPasteValuesAndNumberFormats
    shTarget.Cells(lRow, 9).Value = ngcHrsSum
    shTarget.Cells(lRow, 10).FormulaR1C1 = "=RC[-2]/RC[-1]"

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
