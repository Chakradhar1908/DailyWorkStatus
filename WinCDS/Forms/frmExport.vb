Public Class frmExport
    Public Function ShowExport() As Boolean
        ShowExport = DoShow("export")
    End Function

    Public Function ShowImport() As Boolean
        ShowImport = DoShow("import")
    End Function

    Private Function DoShow(ByVal Mode As String) As Boolean
        Dim C As Control
        Select Case Mode
            Case "import" : C = fraImport
            Case "export" : C = fraType
            Case Else : DevErr("Unknown Mode in frmExport.DoShow: " & Mode)
        End Select

        MoveControl(C, 120, 0, , , True)
        Width = (Width - Me.ClientSize.Width) + picSize.Left
        Height = (Height - Me.ClientSize.Height) + picSize.Top
        Show()

        DoShow = True
    End Function
End Class