Public Class frmProgress2
    Public AltPrg As Object
    Public LockOn As Boolean              ' this does a version modal
    Public PreventLockOn As Boolean
    Private mButtons As MsgBoxStyle

    Public Result As MsgBoxResult
    Public Event Done(ByVal Result As MsgBoxResult, ByRef StayOpen As Boolean)

    Public Sub Progress(Optional ByVal Value As Integer = -1, Optional ByVal Max As Integer = -1, Optional ByVal Cap As String = "#", Optional ByVal DoShow As Boolean = False, Optional ByVal vLockOn As Boolean = True, Optional ByVal vButtons As MsgBoxStyle = 0)
        Dim P As Object
        Result = 0

        If AltPrg Is Nothing Then
            'P = prg    -> prg is name of ucPBar custom control. This control is not completed. After completion of the control, remove the comment.
        Else
            P = AltPrg
        End If

        LockOn = vLockOn
        mButtons = vButtons

        On Error Resume Next
        If Value = -2 Then
            P.Visible = False
            '    PW.Visible = True
            '    PW.Active = True
        Else
            P.Visible = True
            If Not IsIn(Cap, "-", "#", " ", fra.Text) Then fra.Text = Trim(Cap)
            If Max <> -1 Then P.Max = Max
            If Value = -1 Then Value = P.Value + 1
            If Value > P.Max Then Value = P.Max
            P.Value = Value
        End If
        'P.Refresh
        'Me.Refresh
        If DoShow Then
            If Not PreventLockOn Then
                If Value = 1 Or Value = 0 Then
                    Show()
                End If
            End If
            Application.DoEvents()
        End If
    End Sub

End Class