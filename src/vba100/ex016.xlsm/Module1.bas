Option Explicit

Function trimNewLine(str As String) As String
    If Len(str) = 1 Then
        str = ""
    Else
        Dim s, l As Integer
        s = 1: l = Len(str)
        If Left(str, 1) = vbLf Then
            s = 2: l = l - 1
        End If
        If Right(str, 1) = vbLf Then
            l = l - 1
        End If
        str = Mid(str, s, l)
    End If
    trimNewLine = str
End Function

Sub main()
    Dim rng As Range
    Dim cands As Range
    On Error Resume Next
    Set cands = Cells.SpecialCells(xlCellTypeConstants, xlTextValues)
    Msgbox cands.Count
    Err.clear
    If cands Is Nothing Then
        Exit Sub
    End If
    For Each rng In cands
        Dim str As String
        str = rng.Value
        str = Replace(str, vbCrLf, vbLf)
        Do While True
            Dim nxt As String
            nxt = Replace(str, vbLf & vbLf, vbLf)
            If nxt = str Then
                Exit Do
            End If
            str = nxt
        Loop
        rng.Value = trimNewLine(str)
    Next
End Sub
