Imports System.ComponentModel
Public Class frmCashRegisterFND
    Public Cancelled As Boolean

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Cancelled = True
        Hide()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Cancelled = False
        Hide()
    End Sub

    Private Sub txtPrice_Enter(sender As Object, e As EventArgs) Handles txtPrice.Enter
        SelectContents(txtPrice)
    End Sub

    Private Sub cmbVendor_Enter(sender As Object, e As EventArgs) Handles cmbVendor.Enter
        SelectContents(cmbVendor)
    End Sub

    Private Sub txtDesc_Enter(sender As Object, e As EventArgs) Handles txtDesc.Enter
        SelectContents(txtDesc)
    End Sub

    Private Sub Caps(ByVal C As Object)
        Dim SS As Integer, SL As Integer
        On Error Resume Next ' We're assuming they pass a textbox
        SS = C.SelectionStart
        SL = C.SelectionLength
        If C.Text <> UCase(C.Text) Then
            C.Text = UCase(C.Text)
            C.SelectionStart = SS
            C.SelectionLength = SL
        End If
    End Sub

    Private Sub txtPrice_TextChanged(sender As Object, e As EventArgs) Handles txtPrice.TextChanged
        Caps(txtPrice)
    End Sub

    Private Sub txtDesc_TextChanged(sender As Object, e As EventArgs) Handles txtDesc.TextChanged
        Caps(txtDesc)
    End Sub

    Private Sub txtPrice_Validating(sender As Object, e As CancelEventArgs) Handles txtPrice.Validating
        txtPrice.Text = CurrencyFormat(GetPrice(txtPrice.Text))
    End Sub

    Private Sub frmCashRegisterFND_Load(sender As Object, e As EventArgs) Handles Me.Load
        SetButtonImage(cmdOK, 2)
        SetButtonImage(cmdCancel, 3)
        txtPrice.Text = CurrencyFormat(0)
        LoadMfgNamesIntoComboBox(cmbVendor, , False, False)
        cmbVendor.Text = ""
    End Sub

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