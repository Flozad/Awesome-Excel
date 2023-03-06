Sub RemoveDuplicates()
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1:A10")  Change the range as needed
    rng.RemoveDuplicates Columns:=1, Header:=xlNo
End Sub
