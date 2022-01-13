Option Explicit

Sub main()
    Dim wb As Workbook
    Dim cur_wb As Workbook
    Set cur_wb = ThisWorkbook
    Dim txtPath As String
    txtPath = cur_wb.Path & "/" & "log_" & Format(Now(), "yyyymmddhhmmss") & ".txt"
    Dim fnumber As Integer
    fnumber = FreeFile
    Open txtPath For Output As #fnumber
    For each wb In Workbooks
        ' ThisWorkbook closes lastly
        If wb.Name = cur_wb.Name Then
            GoTo Continue
        End If
        Msgbox wb.Name
        Print #fnumber, cur_wb.Path & "/" & wb.Name
        wb.Close saveChanges:=False
Continue:
    Next
    Print #fnumber, cur_wb.Path & "/" & cur_wb.Name
    Close #fnumber
    cur_wb.Close saveChanges:=False
    Application.Quit
End Sub
