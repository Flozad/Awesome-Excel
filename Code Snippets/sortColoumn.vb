Sub SelectColumnA()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Change to the worksheet you want to search
    Dim lastRow As Long
    lastRow = FindLastRow(ws, "A")
    ws.Range("A1:A" & lastRow).Select
End Sub
