Imports Microsoft.Office.Interop.Excel
Public Class FrmProgress3
    Public LockOn As Boolean              ' this does a version modal
    Public PreventLockOn As Boolean
    WithEvents lin1 As Line

    Public Result As MsgBoxResult
    Public Event Done(ByVal Result As MsgBoxResult, ByRef StayOpen As Boolean)

    Public Sub Progress(Optional ByVal Cap As String = "#", Optional ByVal DoShow As Boolean = False, Optional ByVal vLockOn As Boolean = True, Optional ByVal vButtons As MsgBoxStyle = 0)
        Dim P As Object
        Result = 0

        LockOn = vLockOn

        On Error Resume Next
        If Not IsIn(Cap, "-", "#", " ", lblCaption) Then lblCaption.Text = Trim(Cap)

        If DoShow Then
            If Not PreventLockOn Then Show()
            My.Application.DoEvents()
        End If
    End Sub

    Private Sub FrmProgress3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetAlwaysOnTop(Me)
        On Error Resume Next
        'gif.FileName = FXFile("progressbar2.gif")
        gif.Image = Image.FromFile(FXFile("progressbar2.gif"))
    End Sub

    Private Sub FrmProgress3_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        'InitLineBorderForm(lin1, Me, 2)
    End Sub

    Private Sub FrmProgress3_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        'gif.FileName = FXFile("progressbar2.gif")
        gif.Image = Image.FromFile(FXFile("progressbar2.gif"))
    End Sub

    Private Sub FrmProgress3_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        On Error Resume Next
        '  If LockOn And Not PreventLockOn Then Show
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

    Private Sub lblCaption_Click(sender As Object, e As EventArgs) Handles lblCaption.Click
        PreventLockOn = Not PreventLockOn
        SetAlwaysOnTop(Me, False)
    End Sub

    Private Sub FrmProgress3_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        lblCaption_Click(lblCaption, New EventArgs)
    End Sub
End Class