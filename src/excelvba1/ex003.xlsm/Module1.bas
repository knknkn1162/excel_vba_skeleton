Option Explicit

Sub main()
    Dim i As Integer
    ' 偶数行を削除
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
        If i Mod 2 = 0 Then
            Cells(i, 1).Delete Shift:=xlUp
        End If
    Next
    ' １行目に空行
    Rows(1).Insert
    ' A列に空列
    Columns(1).Insert
End Sub