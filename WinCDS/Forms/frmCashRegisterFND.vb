Public Class frmCashRegisterFND
    Public Cancelled As Boolean
    Public Function GetInformation(ByVal Style As String, ByRef Price As Decimal, ByRef Vendor As String, ByRef Desc As String) As Boolean
        lblStyle.Text = Style

        txtPrice.Text = CurrencyFormat(Price)
        cmbVendor.Text = Vendor
        txtDesc.Text = Desc

        'Show vbModal
        ShowDialog()

        Price = GetPrice(txtPrice.Text)
        Vendor = Trim(UCase(cmbVendor.Text))
        Desc = UCase(Trim(Microsoft.VisualBasic.Left(txtDesc.Text, Setup_2Data_DescMaxLen)))

        GetInformation = Not Cancelled
        'Unload Me
        Me.Close()
    End Function

End Class