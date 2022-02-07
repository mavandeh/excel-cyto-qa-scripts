Attribute VB_Name = "PDxFNACH"
Sub RunFNACH()

    Columns("A:H").EntireColumn.Hidden = True
    Range("O:O,P:P").ColumnWidth = 60
    Range("Q:Q,R:R").ColumnWidth = 20
    Columns("S:W").EntireColumn.Hidden = True
    Columns("X:Z").ColumnWidth = 60
    Columns("AA:AM").ColumnWidth = 20
    Cells.RowHeight = 409

End Sub
