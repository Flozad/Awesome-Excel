Sub ConvertRangeToColumn()
'UpdatebyExtendoffice
Dim Range1 As Range, Range2 As Range, Rng As Range
Dim rowIndex As Integer
xTitleId = "KutoolsforExcel"
Set Range1 = Application.Selection
Set Range1 = Application.InputBox("Source Ranges:", xTitleId, Range1.Address, Type:=8)
Set Range2 = Application.InputBox("Convert to (single cell):", xTitleId, Type:=8)
rowIndex = 0
Application.ScreenUpdating = False
For Each Rng In Range1.Rows
    Rng.Copy
    Range2.Offset(rowIndex, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True
    rowIndex = rowIndex + Rng.Columns.Count
Next
Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub
