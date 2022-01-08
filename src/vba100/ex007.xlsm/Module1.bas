Option Explicit

Function formatEraName(str As String, era As String)
    Dim fst As String
    Dim pos As Integer
    fst = Left(era, 1)
    pos = Instr(str, fst)
    If pos <> 0 Then
        If Mid(str, pos, 2) <> era Then
            str = Replace(str, fst, era)
        End If
    End If
    formatEraName = str
End Function

Function formatDate(str As String)
    ' yy/mm/dd のような形に
    str = Replace(str, "元", "1")
    str = Replace(str, " ", "/")
    str = Replace(str, ".", "/")
    str = Replace(str, "-", "/")
    Dim pos As Integer
    ' 年号の正規化
    str = formatEraName(str, "明治")
    str = formatEraName(str, "大正")
    str = formatEraName(str, "昭和")
    str = formatEraName(str, "平成")
    str = formatEraName(str, "令和")
    ' 月のみ -> 日も
    If Right(str, 1) = "月" Then
        str = str & "1日"
    End If
    formatDate = str
End Function

Sub main()
    Dim i As Integer
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Dim m As String
        m = formatDate(Cells(i, 1))
        Dim d As Date
        d = DateValue(m)
        Cells(i, 2) = DateSerial(Year(d), Month(d)+1, 0)
        Cells(i, 2).NumberFormatLocal = "mmdd"
    Next
End Sub
