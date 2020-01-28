Public Class ServiceParts
    Public Enum ServiceForMode
        ServiceMode_ForCustomer = 0
        ServiceMode_ForStock = 1
    End Enum

    Public Sub SelectMode(Optional ByVal Sm As ServiceForMode = ServiceMode_ForCustomer, Optional ByVal ChangeSelectBox As Boolean = False, Optional ByVal ChangeBaseFormMode As Boolean = False)
        Dim Stock As Boolean
        Mode = Sm
        Stock = Not (Sm = ServiceMode_ForCustomer)

        If ChangeBaseFormMode Then
            CreateNewMode = Sm
            Select Case Sm
                Case ServiceMode_ForCustomer : Caption = "Customer Service Parts Order Form"
                Case ServiceMode_ForStock : Caption = "Damaged Stock Parts Order Form"
                Case Else : Caption = "Parts Order Form"
            End Select
        End If

        If ChangeSelectBox Then
            optTagStock = Stock
            optTagCustomer = Not Stock
        End If

        'customer
        fraCustomer.Visible = Not Stock
        '  lblFirstName.Visible = Not Stock
        '  lblLastName.Visible = Not Stock
        '  lblAddress.Visible = Not Stock
        '  lblAddress2.Visible = Not Stock
        '  lblCity.Visible = Not Stock
        '  lblZip.Visible = Not Stock
        '  lblTele1Caption.Visible = Not Stock
        '  lblTele.Visible = Not Stock
        '  lblTele2Caption.Visible = Not Stock
        '  lblTele2.Visible = Not Stock

        ' other non-usable fields for stock mode....
        lblInvoiceNo.Visible = Not Stock
        txtInvoiceNo.Visible = Not Stock
        dteClaimDateCaption.Visible = Not Stock
        dteClaimDate.Visible = Not Stock
        lblSaleNo.Visible = Not Stock
        txtSaleNo.Visible = Not Stock

        ' these are now available for all modes (bfh20050303)
        '  lblRepairCost.Visible = Not Stock
        '  txtRepairCost.Visible = Not Stock
        '  chkPaid.Visible = Not Stock

        '  optCBChargeBack.Visible = Not Stock
        '  optCBDeduct.Visible = Not Stock
        '  optCBCredit.Visible = Not Stock

        '  cmdPrintChargeBack.Visible = Not Stock

        lblServiceOrderNoCaption.Visible = Not Stock
        lblServiceOrderNo.Visible = Not Stock

        lblWhatToDoWStyle.Visible = Stock

        ' whether they can choose store or only see current
        cboStores.Visible = Stock
        txtStoreName.Visible = Not Stock

        ' if it is tagged for customer, we can only do current store
        If Not Stock Then LoadStore 0

' what the 'close window' button says
        cmdMenu.Caption = IIf(Stock, "&Menu", "&Back")

        ' these are always disabled for now... can't switch b/w modes manuallly (hide them??)
        optTagCustomer.Enabled = False
        optTagStock.Enabled = False

        ' make the buttons look OK
        EnableNavigation
    End Sub

End Class