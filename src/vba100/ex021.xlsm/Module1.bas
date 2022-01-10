Option Explicit

Sub main()
    Dim root As String
    root = ThisWorkbook.Path
    Dim fname As String
    Dim backup_dir As String
    backup_dir = root & "/BACKUP"
    fname = Dir(backup_dir & "/")
    Do while fname <> ""
        Dim pos As Integer
        pos = InStr(fname, ".")
        ' len(yyyymmddhhmm)=12
        Dim prev As Date
        prev = DateValue( _
            Mid(fname, pos-12, 4) & "/" & Mid(fname, pos-8, 2) & "/" & Mid(fname, pos-6, 2) _
        )
        ' Trash
        If DateAdd("m", 1, prev) < Date() Then
            ' Msgbox "KIll " & fname
            Kill backup_dir & "/" & fname
        End If
        ' update
        fname = Dir()
    loop
End Sub
