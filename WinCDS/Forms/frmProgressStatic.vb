Public Class frmProgressStatic
    Public AltPrg As Object
    Public LockOn As Boolean              ' this does a version modal
    Public PreventLockOn As Boolean
    Private mButtons As MsgBoxStyle
    Public Result As MsgBoxResult
    Public Event Done(ByVal Result As MsgBoxResult, ByRef StayOpen As Boolean)

    Public Sub ProgressSpin(Optional ByVal Cap As String = "#", Optional ByVal DoShow As Boolean = False, Optional ByVal vLockOn As Boolean = True)
        Const BufferX As Integer = 500
        Const BufferY As Integer = 750

        lbl.Text = Cap                                     ' AutoSize
        'gifSpin.Visible = True

        Width = lbl.Width + BufferX * 2               ' Size Form
        'Height = lbl.Height + gifSpin.Height + BufferY * 2

        'lbl.Move ScaleWidth / 2 - lbl.Width / 2, BufferY
        lbl.Location = New Point(Me.ClientSize.Width / 2 - lbl.Width / 2, BufferY)
        'gifSpin.Move(Width - gifSpin.Width) / 2, lbl.Top + lbl.Height + BufferY / 2
        If DoShow Then
            If Not PreventLockOn Then Show()
            Application.DoEvents()
        End If
    End Sub

    Public Sub Progress(Optional ByVal Cap As String = "#", Optional ByVal DoShow As Boolean = False, Optional ByVal vLockOn As Boolean = True)
        Const BufferX As Integer = 50
        Const BufferY As Integer = 75

        lbl.Text = Cap                                     ' AutoSize

        Width = lbl.Width + BufferX * 2               ' Size Form
        Height = lbl.Height + BufferY * 2
        'gifSpin.Visible = False

        'lbl.Left = ScaleWidth / 2 - lbl.Width / 2     ' Center
        'lbl.Top = ScaleHeight / 2 - lbl.Height / 2
        lbl.Location = New Point(Me.ClientSize.Width / 2 - lbl.Width / 2, Me.ClientSize.Height / 2 - lbl.Height / 2)

        If DoShow Then
            On Error Resume Next
            If Not PreventLockOn Then Show()
            Application.DoEvents()
        End If
    End Sub

    Private Sub frmProgressStatic_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetAlwaysOnTop(Me)
        'On Error Resume Next
        'gifSpin.FileName = FXFile("circleWait.gif")
    End Sub

    Public Sub ProgressClose(Optional ByVal WithEvent As Boolean = False)
        Result = 0
        RaiseEvent Done(0, False)
        'Unload Me
        Me.Close()
    End Sub

    Public Sub Disposee()
        ProgressClose()
    End Sub
End Class