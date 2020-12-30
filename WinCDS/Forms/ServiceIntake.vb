Public Class ServiceIntake
    Public Sub InitForm(ByVal nServiceOrderNumber As Long, ByVal PartsOrderID As Long, ByVal nCBType As Long, ByVal CBAmount As String, ByVal nInvoiceNo As String, ByVal nStore As Long, ByVal nVendor As String)
        If nServiceOrderNumber <> 0 Then
            txtServiceOrderNumber.Visible = True
            lblServiceOrderNumber.Visible = True
            ServiceOrderNumber = nServiceOrderNumber
        Else
            txtServiceOrderNumber.Visible = False
            lblServiceOrderNumber.Visible = False
            ServiceOrderNumber = 0
        End If

        PartsOrderNumber = PartsOrderID
        Store = nStore
        Vendor = nVendor

        CBType = nCBType
        Amount = CBAmount
        InvoiceNo = nInvoiceNo

        Select Case CBType
            Case 0
                txtMode = "Charging Back " & FormatCurrency(Amount)
            Case 1
                txtMode = "Deducting " & FormatCurrency(Amount) & " From Invoice #" & InvoiceNo
            Case 2
                txtMode = "Requesting a Credit of " & FormatCurrency(Amount)
            Case Else
                MsgBox "Unknown letter type!!", vbCritical, "Stop!"
      Unload Me
  End Select

        LoadImages PartsOrderID
End Sub

End Class