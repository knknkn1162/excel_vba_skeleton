Sub main()
    Dim dws As Worksheet
    Set dws = Sheets("データ")
    Dim rng As Range
    Set rng = Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    rng.Columns("D").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],マスタ!C[-3]:C[-1],2,FALSE),"""")"
    rng.Columns("E").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],マスタ!C[-4]:C[-2],3,FALSE),"""")"
    rng.Columns("F").FormulaR1C1 = "=RC[-1]*RC[-3]"
End Sub

