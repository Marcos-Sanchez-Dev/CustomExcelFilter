Sub Button10_Click()
 Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range

    ' Specify the worksheet and table name
    Set ws = ThisWorkbook.Sheets("Main") ' Change "Sheet1" to your sheet name
    Set tbl = ws.ListObjects("Table1") ' Change "Table1" to your table name

    ' Define the criteria range
    Set rng = tbl.ListColumns(1).DataBodyRange ' Change 1 to the column number you want to filter

    ' Clear any existing filters
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Apply the filter to show only yellow cells
    rng.AutoFilter Field:=1, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor
End Sub
Sub Button12_Click()
 Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range

    ' Specify the worksheet and table name
    Set ws = ThisWorkbook.Sheets("Main") ' Change "Sheet1" to your sheet name
    Set tbl = ws.ListObjects("Table1") ' Change "Table1" to your table name

    ' Define the criteria range
    Set rng = tbl.ListColumns(1).DataBodyRange ' Change 1 to the column number you want to filter

    ' Clear any existing filters
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Apply the filter to show only blue cells
     rng.AutoFilter Field:=1, Criteria1:=RGB(0, 112, 192), Operator:=xlFilterCellColor
End Sub

Sub Button15_Click()
Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range

    ' Specify the worksheet and table name
    Set ws = ThisWorkbook.Sheets("Main") ' Change "Sheet1" to your sheet name
    Set tbl = ws.ListObjects("Table1") ' Change "Table1" to your table name

    ' Define the criteria range
    Set rng = tbl.ListColumns(1).DataBodyRange ' Change 1 to the column number you want to filter

    ' Clear any existing filters
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0

    ' Apply the filter to show only red cells
    rng.AutoFilter Field:=1, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterCellColor
End Sub
Sub Button16_Click()
   Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range

    ' Specify the worksheet and table name
    Set ws = ThisWorkbook.Sheets("Main") ' Change "Sheet1" to your sheet name
    Set tbl = ws.ListObjects("Table1") ' Change "Table1" to your table name

    ' Define the criteria range
    Set rng = tbl.ListColumns(1).DataBodyRange ' Change 1 to the column number you want to filter

    ' Clear any existing filters
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0

End Sub

