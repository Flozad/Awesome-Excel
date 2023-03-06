Sub SortData()
    ' Sorts a range of data in ascending order by a specified column
    Dim sortRange As Range
    Set sortRange = ActiveSheet.Range("A1:C10") ' Change the range as needed
    Dim sortColumn As Long
    sortColumn = 2 ' Change the column number as needed
    With sortRange
        .Sort Key1:=.Columns(sortColumn), Order1:=xlAscending, Header:=xlYes
    End With
End Sub