Option Explicit

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).row
        Dim d As Date
        Dim row As Integer
        Dim sales  As Long
        sales = Cells(i, 4)
        d = DateSerial(Cells(i, 1), Cells(i, 2), Cells(i, 3))
        row = Weekday(d, vbMonday) + 1
        Cells(row, 7) = Cells(row, 7) + sales
        Cells(row, 8) = Cells(row, 8) + 1
    Next
    
    For i = 2 To 8
        If Cells(i, 8) = 0 Then
            Cells(i, 9) = 0
        Else
            Cells(i, 9) = Cells(i, 7) / Cells(i, 8)
        End If
    Next
End Sub