Sub InsertRow()
    ' Inserts a new row above the currently selected row
    Selection.EntireRow.Insert Shift:=xlDown
End Sub
