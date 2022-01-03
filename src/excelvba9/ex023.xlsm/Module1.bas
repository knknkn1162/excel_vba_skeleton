Option Explicit
Sub init()
    Dim pos As Integer
    pos = 5
    Cells(pos, 1).CurrentRegion.Offset(1, 0).ClearContents
End Sub
 
Sub main()
    Call init
    Dim ws As Worksheet
    Set ws = Worksheets("練習23")
    Dim d As Date, cond As String
    
    d = ws.Cells(2, 1)
    cond = ws.Cells(2, 2)
    If d = 0 And cond = "" Then
        GoTo EMPTY_EXCEPTION
    End If
     
    Dim pos As Integer
    pos = 5
    With Worksheets("練習23_データ")
        Dim i As Integer
        For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
            Dim flag As Boolean
            flag = True
            If d <> 0 And (Not .Cells(i, 1) = d) Then
                flag = False
            End If
            If cond <> "" And InStr(.Cells(i, 2), cond) = 0 Then
                flag = False
            End If
            If flag = True Then
                ws.Cells(pos, 1).Resize(, 3).Value = .Cells(i, 1).Resize(, 3).Value
                pos = pos + 1
            End If
        Next
    End With
    Exit Sub
EMPTY_EXCEPTION:
    MsgBox "検索条件がありません", vbInformation
End Sub
