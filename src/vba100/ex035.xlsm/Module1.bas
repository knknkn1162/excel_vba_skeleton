Option Explicit

Sub main()
    With Range("B2").CurrentRegion
        .FormatConditions.Delete
    End With

    Dim rng As Range
    Set rng = Range("B2").CurrentRegion
    With Union(rng.Columns(4), rng.Columns(6))
        .FormatConditions.Add _
            Type:=xlCellValue, Operator:=xlLess, Formula1:="90%"
        .FormatConditions(1).Interior.Color = vbRed
        .FormatConditions.Add _
            Type:=xlCellValue, Operator:=xlLess, Formula1:="100%"
        .FormatConditions(2).Font.Color = vbRed
    End With
End Sub
