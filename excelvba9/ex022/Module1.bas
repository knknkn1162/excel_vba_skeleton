Option Explicit

Sub main()
    Dim i As Integer, j As Integer
    Dim master_ws As Worksheet
    Dim ws As Worksheet
    
    ' creeate worksheet if any
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("練習22_回答").Delete
    Err.Clear
    Application.DisplayAlerts = True
    
    Set ws = Worksheets.Add
    With ws
        .Name = "練習22_回答"
        .Cells(1, 1) = "伝票番号"
        .Cells(1, 2) = "商品"
        .Cells(1, 3) = "数量"
        .Cells(1, 4) = "単価"
    End With
    
    Worksheets("練習22").Activate
    Dim pos As Integer
    pos = 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        For j = 0 To 5
            Dim col As Integer
            col = j * 3 + 2
            If Cells(i, col) <> "" Then
                ws.Cells(pos, 1) = Cells(i, 1)
                ws.Range(ws.Cells(pos, 2), ws.Cells(pos, 4)).Value = Range(Cells(i, col), Cells(i, col + 2)).Value
                pos = pos + 1
            End If
        Next
    Next
End Sub