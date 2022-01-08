Option Explicit

Public Function import(xlsm, src_dir)
    Dim xlbook As Workbook
    Dim fso As Object
    Set xlbook = Application.Workbooks.Open(xlsm)
    Dim compos As Object, compo As Object
    Set compos = xlbook.VBProject.VBComponents
    coreUnbind compos
    
    Dim filepath As String
    filepath = Dir(src_dir & "/*")
    Do While filepath <> ""
        compos.import (src_dir & "/" & filepath)
        filepath = Dir()
    Loop
    xlbook.Save
    xlbook.Close
End Function

Public Function unbind(xlsm)
    Dim xlbook As Workbook
    Dim fso As Object
    Set xlbook = Application.Workbooks.Open(xlsm)
    Dim compos As Object, compo As Object
    Set compos = xlbook.VBProject.VBComponents
    coreUnbind compos
    xlbook.Save
    xlbook.Close
End Function

Private Function coreUnbind(compos)
    Const vbext_ct_Document = 100 'VBComponent Type 定数 : 標準モジュール
    ' delete
    Dim compo As Object
    For Each compo In compos
        Select Case compo.Type
            Case vbext_ct_Document
                ' do nothing
            Case Else
                compos.Remove compo
        End Select
    Next compo
End Function
