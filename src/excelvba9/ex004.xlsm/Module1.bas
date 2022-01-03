Option Explicit

Sub main()
    Dim row As Integer
    Dim i As Integer
    row = Cells(Rows.Count, 1).End(xlUp).row
    For i = 2 To row
        Dim cnt As Integer
        cnt = Int(Cells(i, 2) / Cells(i, 3))
        If cnt = 0 Then
            Cells(i, 4) = "x"
        Else
            Cells(i, 4) = cnt
        End If
        Cells(i, 5) = Cells(i, 2) Mod Cells(i, 3)
    Next
End Sub
