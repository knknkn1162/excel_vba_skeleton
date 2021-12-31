Option Explicit

Sub main()
    With Worksheets("Sheet1")
        '.Range("B2") = "日付(" & Left(Range("a1"), 4) & "/" & Mid(Range("a1"), 5, 2) & "/" & Mid(Range("a1"), 7) & ")"
        .Range("B2") = Format(Range("a1"), "'0000/00/00")
    End With
    
    With Worksheets("Sheet2")
        Dim pos As Integer
        Dim str As String
        str = .Range("A1")
        pos = InStr(str, " ")
        .Range("B1") = Left(str, pos - 1)
        .Range("C1") = Mid(str, pos)
    End With
End Sub