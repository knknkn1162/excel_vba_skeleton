Option Explicit

Sub main()
    Dim r As Range
    Dim i As Integer, j As Integer, row As Integer
    ' J 列
    For i = 2 To 4 * 4 + 2 - 1
        Select Case i Mod 4
            Case 3, 0
                Cells(i, 10) = WorksheetFunction.Sum(Range(Cells(i, 3), Cells(i, 9)))
            Case 1
                Cells(i, 10) = Cells(i - 2, 10) / Cells(i - 1, 10)
                Cells(i, 10).NumberFormatLocal = "0.0"
            Case Else
        End Select
    Next
 
    
    For row = 0 To 1
        Dim base As Integer, m As Integer
        base = 18 + row
        m = 3 + row
        For j = 3 To 10
            For i = 0 To 3
                Cells(base, j) = Cells(base, j) + Cells(i * 4 + m, j)
            Next
        Next
    Next
    
        
    'C~I 列の客単価
    For i = 1 To 4
        For j = 3 To 9
            Cells(i * 4 + 1, j) = Cells(i * 4 - 1, j) / Cells(i * 4, j)
            Cells(i * 4 + 1, j).NumberFormatLocal = "0.0"
        Next
    Next
    For j = 3 To 10
        Cells(20, j) = Cells(18, j) / Cells(19, j)
        Cells(20, j).NumberFormatLocal = "0.0"
    Next
    
    '  赤文字
    For j = 3 To 9
        For i = 1 To 4
            If Cells(i * 4 + 1, j) < Cells(i * 4 + 1, 10) Then
                Cells(i * 4 + 1, j).Font.Color = vbRed
            End If
        Next
        If Cells(20, j) < Cells(20, 10) Then
            Cells(20, j).Font.Color = vbRed
        End If
    Next
  
End Sub
