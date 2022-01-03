Option Explicit

Sub init()
    Range("A1").CurrentRegion.Offset(1, 0).ClearContents
End Sub

Sub main()
    Call init
    Dim cur_path As String
    cur_path = ThisWorkbook.FullName
    Dim cur_ws As Worksheet
    Set cur_ws = ActiveSheet
    Dim target_dir As String
    target_dir = Range("A1")
    Dim filename As String
    filename = Dir(target_dir & "/*.xlsm")
    
    Dim row As Integer, col As Integer
    row = 2
    Do While filename <> ""
        col = 1
        Dim path As String
        path = target_dir & "/" & filename

        ' exclude this workbook
        If path = cur_path Then
            GoTo Continue
        End If
        
        With Workbooks.Open(path)
            cur_ws.Cells(row, col) = .Name
            col = col + 1
            Dim ws As Worksheet
            For Each ws In .Worksheets
                cur_ws.Cells(row, col) = ws.Name
                col = col + 1
            Next ws
            .Close SaveChanges:=False
        End With
        row = row + 1
Continue:
        filename = Dir()
    Loop
End Sub
