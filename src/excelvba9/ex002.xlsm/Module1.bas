Option Explicit

Sub init()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 4).Clear
    Next
End Sub

Sub ex1()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 4) = Cells(i, 2) * Cells(i, 3)
    Next
End Sub