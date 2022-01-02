Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 4) = Cells(i, 3) / Cells(i, 2)
        Select Case Cells(i, 4)
            Case Is >= 1.05
                Cells(i, 4).Interior.Color = vbBlue
                Cells(i, 4).Font.Color = vbWhite
            Case Is >= 1
                Cells(i, 4).Font.Color = vbBlue
            Case Is >= 0.95
                Cells(i, 4).Font.Color = vbBlack
            Case Is >= 0.9
                Cells(i, 4).Font.Color = vbRed
            Case Else
                Cells(i, 4).Interior.Color = vbRed
                Cells(i, 4).Font.Color = vbBlack
        End Select
    Next
End Sub