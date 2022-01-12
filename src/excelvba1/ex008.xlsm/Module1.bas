Option Explicit

Function isPathRooted(str As String) As Boolean
    str = Trim(str)
    isPathRooted = False
    If Left(str, 1) = "/" Then
        isPathRooted = True
    End If
    If (Not Application.OperatingSystem Like "*Mac*") And Mid(str, 2, 1) = ":" Then
        isPathRooted = True
    End If
End Function

Sub main()
    Dim fname As String
    Dim root As String
    root = ThisWorkbook.Path
    Dim flag As boolean
    fname = InputBox("フォルダを入力してください" & vbLf & "(相対パスでも可)")
    If Not fname Like "*.xls*" Then
        Msgbox "excelファイルではありません"
        Exit Sub
    End If
    If Not isPathRooted(fname) Then
        fname = root & "/" & fname
    End If
    Dim wb As Workbook
    Dim pos As Integer
    pos = InStrRev(fname, "/")
    Set wb = IIf(Mid(fname, pos+1) = ThisWorkbook.Name, ThisWorkbook, Workbooks.Open(fname))
    
    Dim ws As Worksheet
    For each ws In wb.Worksheets
        ws.ExportAsFixedFormat Type:=xlTypePDF, _
            FileName:= Left(fname, pos-1) & "/" & ws.Name, _
            OpenAfterPublish := False
    Next
End Sub
