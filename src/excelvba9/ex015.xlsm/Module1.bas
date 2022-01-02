Option Explicit
Sub init()
    Dim ws As Worksheet
    Set ws = Worksheets("練習15_回答")
    ' automatically expand the selection to include the entire current region
    ws.Range("A1").CurrentRegion.Offset(1, 1).ClearContents
End Sub

Sub main()
    Call init
    Dim i As Integer
    Dim j As Integer
    Dim dst_ws As Worksheet
    Set dst_ws = Worksheets("練習15_回答")
    With Worksheets("練習15")
        For i = 2 To .Cells(Rows.Count, 1).End(xlUp).row
            Dim row As Integer
            Dim col As Integer
            ' search branch
            For j = 2 To dst_ws.Cells(1, Columns.Count).End(xlToLeft).Column
                If .Cells(i, 1) = dst_ws.Cells(1, j) Then
                     col = j
                     Exit For
                End If
            Next
            ' search goods
            For j = 2 To dst_ws.Cells(Rows.Count, 1).End(xlUp).row
                If .Cells(i, 2) = dst_ws.Cells(j, 1) Then
                    row = j
                    Exit For
                End If
            Next
            dst_ws.Cells(row, col) = dst_ws.Cells(row, col) + .Cells(i, 3)
        Next
    End With
End Sub