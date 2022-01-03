Option Explicit

Sub main()
    Dim i As Integer
    Dim row As Integer
    row = Cells(Rows.Count, 2).End(xlUp).row
    For i = row To 2 Step -1
        Select Case Cells(i, 1)
            Case "I"
                Rows(i).Insert
            Case "D"
                Rows(i).Delete
        End Select
    Next
End Sub
