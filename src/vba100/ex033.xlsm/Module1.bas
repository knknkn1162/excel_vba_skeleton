Sub Macro1()
    Sheets("データ").Select
    For i = 2 To 1001
    Range("D" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-2],マスタ!C[-3]:C[-1],2,FALSE),"""")"
    Range("E" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],マスタ!C[-4]:C[-2],3,FALSE),"""")"
    Range("F" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=RC[-1]*RC[-3]"
    Range("D" & i & ":F" & i).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Next
    Range("A1").Select
End Sub

