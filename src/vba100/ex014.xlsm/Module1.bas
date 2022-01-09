Option Explicit

Sub main()
    Dim ws As WorkSheet
    Dim rng As Range
    For Each ws In WorkSheets
        With ws.Cells
            On Error Resume Next
            Set rng = InterSect(.Cells, .SpecialCells(xlCellTypeFormulas))
            Err.Clear
        End With
        Dim r As Range
        For each r In rng
            r.Value = r.Value
        Next
    Next

    For Each ws In WorkSheets
        If Instr(ws.Name, "社外秘") <> 0 Then
            Application.DisplayAlerts = false
            ws.Delete
            Application.DisplayAlerts = true
        End If
    Next

End Sub
