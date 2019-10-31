Public Class MainMenu4
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Case "newsale"
        If CrippleBug("New Sales") Then Exit Sub
        If Not CheckAccess("Create Sales") Then Exit Sub
        Order = "A"
        'frmSalesList.SafeSalesClear = True
        frmSalesList.SalesCode = ""
        'Unload BillOSale
        BillOSale.Close()
        MainMenu.Hide()
        'BillOSale.HelpContextID = 42000
        'BillOSale.HelpContextID = 42002
        BillOSale.Show()
        'MailCheck.HelpContextID = 42000
        'MailCheck.optTelephone.Value = True
        MailCheck.HidePriorSales = True
        'MailCheck.Show vbModal  ' If this is loaded "vbModal, BillOSale", lockup may occur.
        MailCheck.ShowDialog()
        MailCheck.HidePriorSales = False
        'Unload MailCheck
        MailCheck.Close()
    End Sub
End Class