Sub SortRange()
    ' Sorts a range of cells in the active sheet in ascending order
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1:A10") ' Change the range as needed
    rng.Sort Key1:=rng, Order1:=xlAscending, Header:=xlNo
End Sub
