Option Explicit

Sub main()
    Dim cur_month As Integer
    cur_month = 1
    Dim sum As Long, i As Integer
    ' guard
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = "End"
    i = 2
    Do Until Cells(i, 1) = ""
        Dim tmp As Integer
        If Cells(i, 1) = "End" Then
            tmp = 100
            Cells(i, 1) = ""
        Else
            tmp = month(Cells(i, 1))
        End If
        If tmp <> cur_month Then
            Rows(i).Insert
            Cells(i, 1) = cur_month & "月合計"
            Cells(i, 2) = sum
            Range(Cells(i, 1), Cells(i, 2)).Font.Bold = True
            sum = 0
            i = i + 1
            cur_month = tmp
        End If
        sum = sum + Cells(i, 2)
        i = i + 1
    Loop
End Sub