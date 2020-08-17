Public Class frmTransferReports
    Private mMode As String
    Public Property Mode() As String
        Get
            Mode = mMode
        End Get
        Set(value As String)
            mMode = value
            Arrange
        End Set
    End Property

    Private Sub Arrange(Optional ByVal Working As Boolean = False)
        Me.Cursor = IIf(Working, Cursors.WaitCursor, Cursors.Default)
        cmdPrint0.Enabled = Not Working
        cmdPrint1.Enabled = Not Working
        cmdCancel.Enabled = Not Working

        Enabled = False
        fraPending.Visible = False
        fraPrevious.Visible = False
        Select Case Mode
            Case "Pending"
                fraPending.Visible = True
            Case "Previous"
                fraPrevious.Visible = True
        End Select
        Enabled = True
    End Sub

End Class