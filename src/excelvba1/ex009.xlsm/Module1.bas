Option Explicit

Sub main()
    Dim root As String, backup_dir As String
    root = ThisWorkbook.Path
    backup_dir = root & "/" & "BACKUP"
    On Error Resume Next
    RmDir backup_dir
    Err.Clear
    MkDir backup_dir

    Dim fname As String
    fname = Dir(root & "/*.xls")
    Do While fname <> ""
        If fname = ThisWorkbook.name Then
            GoTo Continue
        End If
        FileCopy source:=root & "/" & fname, _
            destination:=backup_dir & "/" & fname
Continue:
        fname = Dir()
    Loop
End Sub
