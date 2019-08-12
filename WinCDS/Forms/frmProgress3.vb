Public Class FrmProgress3
    Public LockOn As Boolean              ' this does a version modal
    Public PreventLockOn As Boolean

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
            Application.DoEvents()
        End If
    End Sub

    Private Sub FrmProgress3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class