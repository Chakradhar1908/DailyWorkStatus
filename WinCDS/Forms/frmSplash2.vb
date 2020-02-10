Public Class frmSplash2
    Dim PrgValue As Integer, PrgMax As Integer
    Public Sub DoStatus(Optional ByVal Msg As String = "#")
        'On Error Resume Next
        If Msg <> "#" Then Me.lblStatus.Text = Msg
        Refresh()
#If OldMainMenu = 0 Then
        If Not Visible Then
            Show()    ' <-- comment out this line to prevent splash from showing
        End If
#End If
    End Sub

    Public Sub DoClose()
        'Unload Me
        Me.Close()
    End Sub

    Private Sub frmSplash2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetAlwaysOnTop(Me)
        On Error Resume Next
        SetCaptions
        imgBackground.Refresh
    End Sub

    Private Sub SetCaptions()
        lblProgram0.Text = "Loading " & ProgramName & "..."
        lblProgram1.Text = "Copyright " & SoftwareCopyright(True) & "..."
        lblProgram2.Text = "Version: " & SoftwareVersion(False, True)
        lblProgram3.Text = IIf(IsServer, "SERVER", "WORKSTATION")
    End Sub

    Public Sub DoProgress(Optional ByVal Value As Integer = -1, Optional ByVal Max As Integer = -1)
        Dim X As Integer
        On Error Resume Next
        If Max > 0 Then
            If Max < PrgValue Then PrgValue = Max
            PrgMax = Max
        End If

        If Value >= 0 Then
            If Value > PrgMax Then Value = PrgMax
            PrgValue = Value
            'X = picProgress.ScaleWidth * (CDbl(PrgValue) / CDbl(PrgMax))
            X = picProgress.ClientRectangle.Width * (CDbl(PrgValue) / CDbl(PrgMax))
            picProgress.Visible = True
            'picProgress.Cls
            picProgress.Image = Nothing
            'picProgress.FillStyle = vbSolid
            'picProgress.FillColor = vbBlack
            'picProgress.Line(0, 0, X, picProgress.ScaleHeight - 10, vbBlue, B)

            picProgress.Refresh()
        Else
            picProgress.Visible = False
        End If
    End Sub

End Class