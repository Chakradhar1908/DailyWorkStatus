Public Class ServiceParts
    Private Mode As ServiceForMode
    Public CreateNewMode As ServiceForMode
    Dim Store As Integer
    Dim ServiceCallNumber As Integer
    Dim PartsOrderID As Integer
    Public Enum ServiceForMode
        ServiceMode_ForCustomer = 0
        ServiceMode_ForStock = 1
    End Enum

    Public Sub SelectMode(Optional ByVal Sm As ServiceForMode = ServiceForMode.ServiceMode_ForCustomer, Optional ByVal ChangeSelectBox As Boolean = False, Optional ByVal ChangeBaseFormMode As Boolean = False)
        Dim Stock As Boolean
        Mode = Sm
        Stock = Not (Sm = ServiceForMode.ServiceMode_ForCustomer)

        If ChangeBaseFormMode Then
            CreateNewMode = Sm
            Select Case Sm
                Case ServiceForMode.ServiceMode_ForCustomer : Text = "Customer Service Parts Order Form"
                Case ServiceForMode.ServiceMode_ForStock : Text = "Damaged Stock Parts Order Form"
                Case Else : Text = "Parts Order Form"
            End Select
        End If

        If ChangeSelectBox Then
            optTagStock.Checked = Stock
            optTagCustomer.Checked = Not Stock
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
        If Not Stock Then LoadStore(0)

        ' what the 'close window' button says
        cmdMenu.Text = IIf(Stock, "&Menu", "&Back")

        ' these are always disabled for now... can't switch b/w modes manuallly (hide them??)
        optTagCustomer.Enabled = False
        optTagStock.Enabled = False

        ' make the buttons look OK
        EnableNavigation()
    End Sub

    Private Function LoadStore(ByVal StoreNo As Integer) As Boolean  ' Fill in store address.
        If StoreNo = 0 Then StoreNo = StoresSld
        Store = StoreNo
        LoadSoldToAddress(StoreNo)
    End Function

    Public Sub LoadSoldToAddress(ByVal StoreNo As Integer)
        Dim CInfo As StoreInfo

        On Error Resume Next
        cboStores.SelectedIndex = StoreNo - 1
        CInfo = StoreSettings(StoreNo)
        txtStoreName.Text = CInfo.Name
        txtStoreAddress.Text = CInfo.Address
        txtStoreCity.Text = CInfo.City
        txtStorePhone.Text = CInfo.Phone
    End Sub

    Private Sub EnableNavigation()
        Dim ShowNavigate As Boolean, Navigate As Boolean, Search As Boolean
        Dim Mode As Integer

        ShowNavigate = True
        cmdMoveFirst.Visible = ShowNavigate
        cmdMovePrevious.Visible = ShowNavigate
        cmdMoveNext.Visible = ShowNavigate
        cmdMoveLast.Visible = ShowNavigate
        cmdMoveSearch.Visible = ShowNavigate
        lblMoveRecords.Visible = ShowNavigate


        If ServiceCallNumber <> 0 Then
            Mode = IIf(PartsOrderID <> 0, 1, 2)
        Else
            Mode = IIf(PartsOrderID <> 0, 1, 3)
        End If

        Select Case Mode
            Case 1  ' show all
                cmdMoveFirst.Width = 375
                SetButtonImage(cmdMoveFirst, "previous")
                cmdMoveFirst.Text = ""
                Navigate = True : Search = True
            Case 2  ' show 1
                cmdMoveFirst.Width = 2655
                cmdMoveFirst.Text = "Browse Records"
                'cmdMoveFirst.Picture = Nothing
                cmdMoveFirst.Image = Nothing
                Navigate = False : Search = False
            Case 3  ' show browse and search
                cmdMoveFirst.Width = 1455
                cmdMoveFirst.Text = "Browse Records"
                'cmdMoveFirst.Picture = Nothing
                cmdMoveFirst.Image = Nothing
                Navigate = False : Search = True
        End Select

        cmdMoveLast.Enabled = Navigate
        cmdMoveNext.Enabled = Navigate
        cmdMovePrevious.Enabled = Navigate
        cmdMoveSearch.Enabled = Search
    End Sub
End Class