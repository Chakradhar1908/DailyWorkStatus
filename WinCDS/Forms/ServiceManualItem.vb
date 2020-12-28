Public Class ServiceManualItem
    Public Cancelled As Boolean

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Cancelled = True
        Hide()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Hide()
    End Sub

    Private Sub ClearForm()
        Cancelled = False
        txtSaleNo.Text = ""
        txtStyle.Text = ""
        txtQuantity.Text = "1"
        txtDesc.Text = ""
        dtpDelDate.Value = Today
        cboVendor.Text = ""
    End Sub

    Private Sub ServiceManualItem_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        ColorDatePicker(dtpDelDate)
        LoadMfgNamesIntoComboBox(cboVendor, , , True)
        ClearForm()
    End Sub

    'Private Sub Modify(ByRef txt As TextBox)
    Private Sub Modify(ByRef txt As Object)
        Dim S As Integer
        On Error Resume Next
        S = txt.SelectionStart
        txt.Text = UCase(txt.Text)
        txt.SelectionStart = S
    End Sub

    Private Sub cboVendor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboVendor.SelectedIndexChanged
        Modify(cboVendor)
    End Sub

    Private Sub txtDesc_TextChanged(sender As Object, e As EventArgs) Handles txtDesc.TextChanged
        Modify(txtDesc)
    End Sub

    Private Sub txtQuantity_Leave(sender As Object, e As EventArgs) Handles txtQuantity.Leave
        txtQuantity.Text = Val(txtQuantity.Text)
    End Sub

    Private Sub txtSaleNo_TextChanged(sender As Object, e As EventArgs) Handles txtSaleNo.TextChanged
        Modify(txtSaleNo)
    End Sub

    Private Sub txtStyle_TextChanged(sender As Object, e As EventArgs) Handles txtStyle.TextChanged
        Modify(txtStyle)
    End Sub
End Class