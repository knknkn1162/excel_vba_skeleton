Option Explicit

Sub main()
    Dim db_ws As Worksheet
    Set db_ws = Worksheets("練習16_マスタ")
    With Worksheets("練習16")
        Dim i As Integer
        Dim row As Integer, sum As Long
        sum = 0
        row = .Cells(Rows.Count, 1).End(xlUp).row
        For i = 2 To row
            Dim cnt As Integer
            cnt = WorksheetFunction.CountIf(db_ws.Columns(1), .Cells(i, 2))
            If cnt > 0 Then
                Dim pos As Integer
                pos = WorksheetFunction.Match(.Cells(i, 2), db_ws.Columns(1), False)
                .Cells(i, 3) = db_ws.Cells(pos, 2)
                .Cells(i, 4) = db_ws.Cells(pos, 3)
                .Cells(i, 6) = .Cells(i, 4) * .Cells(i, 5)
                sum = sum + .Cells(i, 6)
            End If
        Next
        .Cells(row + 1, 6) = sum
    End With
End Sub