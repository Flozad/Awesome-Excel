Sub CopyConditionalFormatting()
    ' Copies the conditional formatting from one range of cells to another
    Dim sourceRange As Range
    Dim targetRange As Range
    Set sourceRange = ActiveSheet.Range("A1:A10") ' Change the source range as needed
    Set targetRange = ActiveSheet.Range("B1:B10") ' Change the target range as needed
    sourceRange.Copy
    targetRange.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub
