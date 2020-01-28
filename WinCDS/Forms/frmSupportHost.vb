Public Class frmSupportHost
    Public Sub Listen(Optional ByVal DoShow As Boolean = True)
        sbc.Listen
        sbc.Visible = False
        If DoShow Then Show()
    End Sub

End Class