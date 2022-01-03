Option Explicit

Sub main()
    Dim i As Integer
    Dim dst_ws As Worksheet, src_ws As Worksheet
    Set src_ws = Worksheets("ブック一覧")
   
    ' delete worksheet("シート一覧") if any
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets("シート一覧").Delete
    Err.Clear
    Application.DisplayAlerts = True
    
    ' create worksheet
    Set dst_ws = Worksheets.Add
    With dst_ws
        .Name = "シート一覧"
        .Range("A1") = "ブック名"
        .Range("B1") = "シート名"
    End With
    src_ws.Activate
    Dim row As Integer
    row = 2
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).row
        Application.ScreenUpdating = False
        With Workbooks.Open(src_ws.Cells(i, 1) & "/" & src_ws.Cells(i, 2))
            Dim ws As Worksheet
            dst_ws.Cells(row, 1) = .Name
            For Each ws In .Worksheets
                dst_ws.Cells(row, 2) = ws.Name
                row = row + 1
            Next ws
            .Close savechanges:=False
        End With
        Application.ScreenUpdating = True
    Next
End Sub
