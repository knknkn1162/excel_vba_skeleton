Option Explicit

Sub main()
    Dim rng As Range
    For Each rng In Range("A1").CurrentRegion
        If rng.Address <> rng.MergeCells(1).Address Then
            Continue
        End If
        If rng.MergeCells Then
            rng.AddComment "セル結合されています"
            ' rng.Comment.Visible = True
        End If
Continue:
    Next
End Sub
