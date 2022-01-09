Option Explicit

Sub main()
    Dim d As Date
    Dim s As WorkSheet, heads As WorkSheet, prevs As WorkSheet
    Set heads = WorkSheets.Add(Before:=WorkSheets(1))
    heads.Name = "guard"
    Set prevs = heads
    d = DateSerial(2020, 4, 1)
    Do
        Set s = WorkSheets(Format(d, "yyyy年mm月"))
        s.Move After:=prevs
        Set prevs = s
        d = DateValue(DateAdd("m", 1, d))
    Loop Until d = DateSerial(2021, 4, 1)
    heads.Delete
End Sub
