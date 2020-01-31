Public Class frmUpgradeNotify
    Public Sub Notify(ByVal Msg As String, Optional ByVal mCaption As String = "", Optional ByVal Details As String = "", Optional ByVal AllowChangeLog As Boolean = True, Optional ByVal Modal As Boolean = False)
        'If Msg = "" Then Unload Me: Exit Sub
        If Msg = "" Then Me.Close() : Exit Sub
        Text = IIf(mCaption = "", "Attention", mCaption)
        lbl.Text = Msg

        If Details = "" And AllowChangeLog Then
            Details = ReadFile(AppFolder() & "ChangeLog.txt")
        End If

        lblDetails.Text = Details

        'Show IIf(Modal, 1, 0)
        If Modal = 1 Then
            ShowDialog()
        Else
            Show()
        End If
        SetAlwaysOnTop(Me)
    End Sub

End Class