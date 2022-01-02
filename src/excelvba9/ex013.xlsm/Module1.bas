Option Explicit

Sub main()
     Dim i As Integer
     For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Dim pos As Integer
        pos = InStr(Cells(i, 1), "(") ' pref(hiragana)
        Dim pre_pref As String
        pre_pref = Left(Cells(i, 1), pos - 1) ' pref(kanji)
        Dim pos2 As Integer
        pos2 = InStr(Cells(i, 2), "(") ' city(hiragana)
        If InStr(Mid(Cells(i, 1), pos), Mid(Cells(i, 2), pos2 + 1)) > 0 Then
            ' If match
            Cells(i, 3) = pre_pref
        Else
            Cells(i, 3) = pre_pref & "(" & Left(Cells(i, 2), pos2 - 1) & ")"
        End If
     Next
End Sub