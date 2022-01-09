Option Explicit

Sub main()
    Dim rng As Range, r As Range
    For Each rng In Range("A1").CurrentRegion
        If rng.MergeCells = False Then
            goto Continue
        End If
        Dim unit As Range
        Set unit = rng.MergeArea
        rng.Unmerge
        Dim value As Long
        value = Int(rng.Value / unit.CountLarge)
        Dim m As Long
        m = rng.Value Mod unit.CountLarge
        unit.Value = value
        unit.Resize(m).Value = value + 1
Continue:
    Next
End Sub
