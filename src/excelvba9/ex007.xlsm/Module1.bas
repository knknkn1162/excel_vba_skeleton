Option Explicit

Sub main()
    Dim i As Integer
    Dim row As Integer
    row = Cells(Rows.Count, 1).End(xlUp).row
    Dim sum As Long
    sum = 0
    For i = 2 To row
        Cells(i, 4) = Cells(i, 2) * Cells(i, 3)
        sum = sum + Cells(i, 4)
    Next
    Cells(row + 1, 4) = sum
    MsgBox sum & vbLf & sum / (row - 1), vbOKOnly, "結果"
End Sub
