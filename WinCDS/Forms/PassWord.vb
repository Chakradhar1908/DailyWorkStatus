Public Class PassWord
    ' Time to rewrite this form..
    ' It needs a (optional) username box,
    ' a (required) password box,
    ' and a (optional) confirmation box.
    '  Also an Apply button in all cases..

    ' The form will be used to collect passwords for inline validation.
    ' It will also be used to change an account password.

    Dim Mode as integer      ' 1 = Get, 2 = Change.
    Dim PWEntry As String

    Public Function GetPassword(Optional ByRef ParentForm As Form = Nothing, Optional ByVal Reason As String = "", Optional ByVal Zone As String = "") As String
        ' Pop up the form, with (username and) password box visible.
        ' On Apply, return the result and unload the form?
        Mode = 1
        Height = 2140

        lblName.Visible = False
        txtName.Visible = False
        lblOldPassword.Visible = False
        txtOldPassword.Visible = False
        lblConfirm.Visible = False
        txtConfirm.Visible = False

        lblPassword.Top = lblName.Top
        txtPassword.Top = txtName.Top
        'txtPassword.ToolTipText = "Enter password for [" & Zone & "]"
        cmdApply.Top = txtPassword.Top + txtPassword.Height + 120
        cmdCancel.Top = cmdApply.Top

        Dim fontsize As Font = lblPassword.Font
        'lblPassword.FontSize = 14
        lblPassword.Font = New Font(fontsize, 14)
        'lblPassword.FontBold = True
        lblPassword.Font = New Font(fontsize, FontStyle.Bold)
        lblPassword.Text = IIf(Reason = "", "Password:", Reason)
        If Len(Reason) > 12 Then
            'lblPassword.FontSize = 11
            lblPassword.Font = New Font(fontsize, 11)
            'lblPassword.FontBold = False
            lblPassword.Font = New Font(fontsize, FontStyle.Bold)
        End If

        Height = cmdApply.Top + cmdApply.Height + 480

        'Show vbModal, ParentForm
        Me.ShowDialog(ParentForm)
        GetPassword = PWEntry

        lblPassword.Text = "New Password:"

        PWEntry = ""

    End Function

End Class