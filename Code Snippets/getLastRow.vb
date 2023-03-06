Function GetLastRow(ByVal column As String) As Long
    ' Returns the last used row number in the specified column
    GetLastRow = ActiveSheet.Cells(Rows.Count, column).End(xlUp).Row
End Function
