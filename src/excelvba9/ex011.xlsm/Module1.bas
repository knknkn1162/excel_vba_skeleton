Option Explicit

 Sub main()
    Range("A1:B6").Copy Destination:=Range("d1")
    ' Range("A1:B6"  ).Copy
    ' Range("G1").PasteSpecial Paste:=xlPasteValues
    Range("G1:H6").Value = Range("A1:B6").Value
 End Sub