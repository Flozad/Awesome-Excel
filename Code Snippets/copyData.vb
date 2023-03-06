Sub CopyData()
    ' Copies a range of data to the clipboard
    Dim copyRange As Range
    Set copyRange = ActiveSheet.Range("A1:C10") ' Change the range as needed
    copyRange.Copy
End Sub
