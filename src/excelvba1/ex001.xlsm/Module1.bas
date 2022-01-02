Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 3) = Cells(i, 1) * Cells(i, 2)
    Next
End Sub