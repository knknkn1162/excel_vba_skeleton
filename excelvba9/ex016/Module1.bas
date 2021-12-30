Option Explicit

Sub main()
    Dim db_ws As Worksheet
    Set db_ws = Worksheets("練習16_マスタ")
    Dim db_row As Integer
    db_row = db_ws.Cells(Rows.Count, 1).End(xlUp).row
    With Worksheets("練習16")
        Dim i As Integer, j As Integer
        Dim row As Integer, sum As Long
        sum = 0
        row = .Cells(Rows.Count, 1).End(xlUp).row
        For i = 2 To row
            For j = 2 To db_row
                If .Cells(i, 2) = db_ws.Cells(j, 1) Then
                    .Cells(i, 3) = db_ws.Cells(j, 2)
                    .Cells(i, 4) = db_ws.Cells(j, 3)
                    Exit For
                End If
            Next
             .Cells(i, 6) = .Cells(i, 4) * .Cells(i, 5)
             sum = sum + .Cells(i, 6)
        Next
        .Cells(row + 1, 6) = sum
    End With
End Sub