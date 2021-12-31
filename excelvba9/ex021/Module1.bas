Option Explicit

Sub main()
    Dim i As Integer, pos As Integer
    pos = 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Dim v As Integer
        On Error Resume Next
        v = WorksheetFunction.Match(Cells(i, 1), Columns(4), 0)
        If Err.Number <> 0 Then
            Err.Clear
            v = pos
            pos = pos + 1
        End If
        Cells(v, 4) = Cells(i, 1)
        Cells(v, 5) = Cells(v, 5) + Cells(i, 2)
    Next
End Sub


Sub 練習問題21_2()
    Dim i As Long
    Range("D1").CurrentRegion.Offset(1).ClearContents '出力範囲の消去
    Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CopyToRange:=Range("D1"), _
        Unique:=True
    For i = 2 To Cells(Rows.Count, 4).End(xlUp).Row
        Cells(i, 5) = WorksheetFunction.SumIf(Columns(1), Cells(i, 4), Columns(2))
    Next
End Sub
