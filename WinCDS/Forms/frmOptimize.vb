Public Class frmOptimize
    Public Network As TSPNetwork

    Private Sub frmOptimize_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Move 0, Screen.Height - Height - 1000
        Me.Location = New Point(0, Screen.PrimaryScreen.Bounds.Height - Me.Height - 100)
        SetAlwaysOnTop(Me)
        'HelpContextID = 59650
    End Sub
End Class