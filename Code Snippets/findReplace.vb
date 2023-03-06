Sub FindReplace()
    ' Finds and replaces a specific value in a range of cells in the active sheet
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1:A10") ' Change the range as needed
    rng.Replace What:="OldValue", Replacement:="NewValue", LookAt:=xlWhole, _
                MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
