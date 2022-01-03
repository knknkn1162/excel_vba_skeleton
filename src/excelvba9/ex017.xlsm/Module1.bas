Option Explicit

Sub init()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Name = "練習17_回答" Then
            ws.Delete
            Exit For
        End If
    Next ws
    Application.DisplayAlerts = True
End Sub

Sub main()
    Call init
    Dim ws As Worksheet
    Dim new_ws As Worksheet
    Dim cur_ws As Worksheet
    Set cur_ws = Worksheets("練習17")
    Set new_ws = Sheets.Add()
    ' initialize
    With new_ws
        .Name = "練習17_回答"
        .Cells(1, 1) = "ブック名"
        .Cells(1, 2) = "シート名"
    End With
    
    Dim pos As Integer
    pos = 2
    Dim ref_wb As Workbook
    Dim i As Integer
    For i = 2 To cur_ws.Cells(Rows.Count, 1).End(xlUp).Row
        With Workbooks.Open(cur_ws.Cells(i, 1) & "/" & cur_ws.Cells(i, 2))
            new_ws.Cells(pos, 1) = .Name
            For Each ws In .Sheets
                new_ws.Cells(pos, 2) = ws.Name
                pos = pos + 1
            Next ws
            .Close SaveChanges:=False
        End With
    Next
 End Sub
