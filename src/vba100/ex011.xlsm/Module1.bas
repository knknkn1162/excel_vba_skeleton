Option Explicit

Sub main()
    Dim flag As Variant
    flag = Range("A1").CurrentRegion.MergeCells 
    If isNull(flag) Or flag = True Then
        Msgbox "セル結合されています"
    End If
End Sub
