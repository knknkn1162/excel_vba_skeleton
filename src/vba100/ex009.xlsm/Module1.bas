Option Explicit

Sub main()

    Dim ws As Worksheet
    Set ws = Worksheets("成績表")
    Application.DisplayAlerts = False
    On Error Resume Next
    WorkSheets("合格者").Delete
    Err.Clear
    Application.DisplayAlerts = True

    Dim dst_ws As Worksheet
    Set dst_ws = Worksheets.Add
    With dst_ws
        dst_ws.Name = "合格者"
    End With

    Dim i As Integer, pos As Integer
    pos = 1
    With ws
        For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
            If .Cells(i,7) = "合格" Then
                dst_ws.Cells(pos, 1) = .Cells(i, 1)
                pos = pos + 1
            End If
        Next
    End With
End Sub
