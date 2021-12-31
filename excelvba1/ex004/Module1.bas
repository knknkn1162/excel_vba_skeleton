Option Explicit

Sub init()
    Range("B2").CurrentRegion.Clear
End Sub

Sub main()
    With Range("B2:C2")
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
    With Range("D2:E2")
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    With Range("B2:E3")
        .Interior.Color = vbBlue
        .Font.Color = vbWhite
    End With
    With Range("B2").CurrentRegion
        .Borders.LineStyle = xlContinuous
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Offset(1, 0).NumberFormatLocal = "#,##0"
    End With

End Sub