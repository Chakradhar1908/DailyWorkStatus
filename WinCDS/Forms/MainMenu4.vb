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

    Private Sub MainMenu4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'NOTE: In vb6, for image control(imgPicture) assigned datasource as datacontrl and datafied as "Picture" column(code is in mod2DataPictures modules ->GetDatabasePicture function).
        'Replacement for it in vb.net is the below line. This code line is not in vb6. Values are directly assigned in the design time properties window of imgPicture image control.

        'imgPicture.DataBindings.Clear()  NOTE: REMOVE THIS COMMENTE IF imgPicture.DataBindings.Add will expect Clear first before Add.
        imgPicture.DataBindings.Add("Image", datPicture, "Picture")
    End Sub
End Class