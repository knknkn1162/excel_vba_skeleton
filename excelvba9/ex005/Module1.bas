Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 4) = Cells(i, 3) / Cells(i, 2)
        Dim rank As String
        Select Case Cells(i, 4)
            Case Is >= 1.05
                rank = "S"
            Case Is >= 1
                rank = "A"
            Case Is >= 0.95
                rank = "B"
            Case Is >= 0.9
                rank = "C"
            Case Else
                rank = "D"
        End Select
        Cells(i, 5) = rank
    Next
End Sub