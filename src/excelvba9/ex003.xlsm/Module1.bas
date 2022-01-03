Option Explicit
Sub init()
    Dim row, col As Integer
    For row = 2 To Cells(Rows.Count, 1).End(xlUp).row
        For col = 2 To Cells(1, Columns.Count).End(xlToLeft).Column
            Cells(row, col).Clear
        Next
    Next
End Sub


Sub main()
    Dim row, col As Integer
    For row = 2 To Cells(Rows.Count, 1).End(xlUp).row
        For col = 2 To Cells(1, Columns.Count).End(xlToLeft).Column
            Cells(row, col) = Cells(row, 1) * Cells(1, col)
        Next
    Next
End Sub
