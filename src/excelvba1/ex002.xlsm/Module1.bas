Option Explicit

 Sub main()
    Dim threshold As Integer
    threshold = InputBox("input number")
    Dim i As Integer
    Dim sum As Long
    sum = 0
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) >= threshold Then
            sum = sum + Cells(i, 1)
        End If
    Next
    MsgBox " 合計: " & sum
 End Sub