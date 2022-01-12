Option Explicit

Sub main()
    Dim rng As Range
    With Range("A1").CurrentRegion
        Set rng = Intersect(.Cells, .Offset(1,1))
    End With
    rng.Clear
    rng.NumberFormatLocal = "@"
End Sub
