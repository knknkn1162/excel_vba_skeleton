Option Explicit

Sub main()
    'Dim i As Integer, j As Integer
    'For i = 3 To Cells(Rows.Count, 2).End(xlUp).Row
    '    For j = 3 To Cells(2, Columns.Count).End(xlToLeft).Column
    '        Cells(i, j) = "=" & Cells(i, 2) & " * " & Cells(2, j)
    '    Next
    'Next
    Dim size As Integer
    size = 10
    Range("C3").Resize(size, size) = "=C$2 * $B3"

    Cells(2, 2).CurrentRegion.Copy
    With Worksheets("sheet2").Cells(2, 2)
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteValues
    End With
    Application.CutCopyMode = False
End Sub