Option Explicit

Sub main()
    Dim rng As Range
    For Each rng In Range("A1").CurrentRegion
        If rng.MergeCells Then
            rng.AddComment = "セル結合されています"
        End If
    Next
End Sub
