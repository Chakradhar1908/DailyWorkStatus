Public Class Practice
    Private mConvPrg As ProgressBar
    Public ReadOnly Property ConversionPrg As ProgressBar
        Get
            ConversionPrg = mConvPrg
        End Get
    End Property

    Public Sub StartupFailure(Optional ByVal OnPurpose As Boolean = False)
        Dim S As String, M As String, N As String
        '  fraControls.Visible = False
        cmdConvertOld.Visible = False
        'cmdWinCDSOnly.Value = False

        '  cmdFunctions.Visible = False
        lblLoc.Visible = False
        updLoc.Visible = False
        txtLoc.Visible = False
        fraStartupCrash.Visible = True
        'fraStartupCrash.ZOrder 0
        fraStartupCrash.BringToFront()
        Show()
        S = ""
        N = vbCrLf
        S = S & M & ""
        S = S & M & "Your software has failed to load correctly."
        S = S & N & "This screen exists to help you attempt to alleviate this error."
        If Not OnPurpose Then MessageBox.Show(S)
    End Sub

End Class