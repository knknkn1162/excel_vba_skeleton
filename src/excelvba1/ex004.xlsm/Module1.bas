Option Explicit

Sub init()
    Range("B2").CurrentRegion.Clear
End Sub

Sub main()
    Range("B2:C2").Merge
    Range("D2:E2").Merge
    
    With Range("B2:E3")
        .HorizontalAlignment = xlCenter
        .Interior.Color = vbBlue
        .Font.Color = vbWhite
    End With
    With Range("B2").CurrentRegion
        .Borders.LineStyle = xlContinuous
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Offset(1, 0).NumberFormatLocal = "#,##0"
    End With

End Sub
