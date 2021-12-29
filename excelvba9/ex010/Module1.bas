Option Explicit

Sub main()
    Dim i As Integer
    Dim row As Integer
    row = Cells(Rows.Count, 1).End(xlUp).row
    For i = 2 To row
        Cells(i, 4) = Cells(i, 2) / Cells(i, 3)
    Next
    Range(Cells(1, 1), Cells(row, 4)).Borders.LineStyle = xlDot
    Range(Cells(1, 1), Cells(1, 4)).BorderAround LineStyle:=xlContinuous
    Range(Cells(1, 1), Cells(row, 1)).BorderAround LineStyle:=xlContinuous
    Range(Cells(1, 1), Cells(row, 4)).BorderAround Weight:=xlThick
End Sub