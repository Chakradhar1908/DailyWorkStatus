Public Class ServiceParts
    Private Mode As ServiceForMode
    Public CreateNewMode As ServiceForMode
    Dim Store As Integer
    Dim ServiceCallNumber As Integer
    Dim PartsOrderID As Integer
    Dim MarginLine As Integer
    Public Vendor As String

    Public Enum ServiceForMode
        ServiceMode_ForCustomer = 0
        ServiceMode_ForStock = 1
    End Enum

    Public Function SetOwner(ByRef frmOwner As Form) As Boolean
        If Not Owner Is Nothing Then
            SetOwner = False
            MessageBox.Show("Error setting owner of ServiceParts form.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Function
        End If
        Owner = frmOwner
        ' Also affect navigation buttons.
        ' On Unload, notify owner.
    End Function

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

    Public Sub LoadInfoFromMarginLine(Optional ByVal ML As Integer = -1, Optional ByVal HideSaleNo As Boolean = False)
        Dim Margin As CGrossMargin

        Margin = New CGrossMargin

        If ML <> -1 Then MarginLine = ML

        If Margin.Load(CStr(MarginLine), "#MarginLine") Then
            txtStyleNo.Text = Margin.Style
            txtDescription.Text = Margin.Desc
            txtSaleNo.Text = Margin.SaleNo
            GetInvoiceInfoFromSaleNo(Margin.SaleNo)
            txtSaleNo.Tag = IIf(txtSaleNo.Text = "", "", "VALID")
            LoadVendor(Margin.Vendor)
            '    Notes_Text.Text = "Margin Line " & AddedML  ' removed:  bfh20050218
        Else
            MessageBox.Show("Error: Can't load GrossMargin item.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            MarginLine = 0
        End If

        lblMarginLine.Text = MarginLine

        If MarginLine = 0 Then
            txtStyleNo.Text = ""
            txtDescription.Text = ""
            txtSaleNo.Text = ""
        End If
        If HideSaleNo Then
            txtSaleNo.Visible = False
            lblSaleNo.Visible = False
        End If

        DisposeDA(Margin)
    End Sub

    Public Function LoadServiceCall(ByVal sC As Integer) As Boolean
        ' Load as much as we can from Service Call Number (vendor, cust, etc).
        ServiceCallNumber = sC
        lblServiceOrderNo.Text = sC
        Dim nSC As clsServiceOrder
        Dim MailRec As clsMailRec

        nSC = New clsServiceOrder
        If nSC.Load(CStr(sC), "#ServiceOrderNo") Then
            ' Fill in necessary SC-related information.
            ' Customer name and address
            LoadServiceCall = True
            ' Look up mail record for this information:
            MailRec = New clsMailRec
            If MailRec.Load(nSC.MailIndex, "#Index") Then
                lblFirstName.Text = MailRec.First
                lblLastName.Text = MailRec.Last
                lblAddress.Text = MailRec.Address
                lblAddress2.Text = MailRec.AddAddress
                lblCity.Text = MailRec.City
                lblZip.Text = MailRec.Zip
                lblTele.Text = DressAni(CleanAni(MailRec.Tele))
                lblTele2.Text = DressAni(CleanAni(MailRec.Tele2))
                Dim Mail2 As MailNew2
                modMail.Mail2_GetAtIndex(MailRec.Index, Mail2)
                lblTele3.Text = Mail2.Tele3
                UpdateTelephoneLabels(MailRec.PhoneLabel1, MailRec.PhoneLabel2, Mail2.PhoneLabel3)
            Else ' Can't load mail record.
                lblLastName.Text = "Missing Customer Information"
                UpdateTelephoneLabels("", "", "")
            End If
        Else
            ' Can't load service call!
            ' Calling functions will handle error messages.
            LoadServiceCall = False
        End If

        ' Some operations are affected when ordering parts for a customer:
        '  Parts list is limited to parts the customer has purchased.
        '  When displaying parts in AddOnAcc, omit parts on current SO.
        '  Navigation is only possible within the SO/Customer.
        '   Different vendors are possible, so many PO/SO can happen.
        EnableNavigation()

        DisposeDA(nSC, MailRec)
    End Function

    Public Function LoadRelativePartsOrder(ByVal Dir As Integer, Optional ByVal Max As Boolean = False, Optional ByVal RestrictToCurrentServiceCall As Boolean = True) As Boolean
        Dim SQL As String, BaseRestrict As String, DirS As String, DirP As String
        Dim RS As ADODB.Recordset, NewID As Integer

        If Dir = 0 Then Exit Function
        BaseRestrict = "WHERE (TRUE=TRUE)" ' allows adding additional " AND ..." clauses w/o checks
        If CreateNewMode = ServiceForMode.ServiceMode_ForCustomer And ServiceCallNumber <> 0 And RestrictToCurrentServiceCall Then
            BaseRestrict = BaseRestrict & " AND (ServiceOrderNo = " & ServiceCallNumber & ")"
        End If

        If Max Then
            DirS = ""
            DirP = IIf(Dir < 0, " ASC", " DESC")
        Else
            DirS = " AND (ServicePartsOrderNo" &
           IIf(Dir < 0, "<", ">") &
           PartsOrderID &
          ")"
            DirP = IIf(Dir > 0, " ASC", " DESC")
        End If

        SQL = "SELECT TOP 1 ServicePartsOrderNo FROM ServicePartsOrder " & BaseRestrict & DirS &
        " ORDER BY ServicePartsOrderNo" & DirP
        RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())

        On Error GoTo NoID
        NewID = 0
        NewID = RS("ServicePartsOrderNo").Value
        If NewID <> 0 Then
            ClearServiceCall(True)
            LoadPartsOrder(NewID)
        End If
        LoadRelativePartsOrder = True
        DisposeDA(RS)
NoID:
    End Function

    Public Sub LoadPartsOrder(ByVal PO As Integer)
        ' Load Parts Order (+items, service call, etc)

        ' The PO has already been saved.  We'll need to know if it's a Customer/Stock order.
        ' We'll also need all the details to populate the form..
        Dim cParts As clsServicePartsOrder
        cParts = New clsServicePartsOrder
        If cParts.Load(CStr(PO), "#ServicePartsOrderNo") Then ' We've got the record, populate the form.
            PartsOrderID = cParts.ServicePartsOrderNo
            LoadStore(cParts.Store)
            MarginLine = cParts.MarginLine
            If MarginLine <> 0 Then
                LoadInfoFromMarginLine() ' Also load the actual parts..
            Else
                txtStyleNo.Text = cParts.Style
                txtDescription.Text = cParts.Desc
            End If

            If Not Len(txtInvoiceNo.Text) And Len(cParts.InvoiceNo) Then txtInvoiceNo.Text = cParts.InvoiceNo
            dteClaimDate.Value = cParts.InvoiceDate

            lblPartsOrderNo.Text = PO
            lblClaimDate.Text = Format(cParts.DateOfClaim, "mm/dd/yy")

            Vendor = cParts.Vendor
            txtVendorName.Text = cParts.Vendor
            txtVendorAddress.Text = cParts.VendorAddress
            txtVendorCity.Text = cParts.VendorCity
            txtVendorTele.Text = cParts.VendorTele

            If cParts.ServiceOrderNo <> 0 Then
                SelectMode(ServiceForMode.ServiceMode_ForCustomer, True)
                LoadServiceCall(cParts.ServiceOrderNo)
            Else
                SelectMode(ServiceForMode.ServiceMode_ForStock, True)
            End If

            SelectStatus(cParts.Status)

            SelectChargeBackOption(cParts.ChargeBackType)
            txtRepairCost.Text = FormatCurrency(cParts.ChargeBackAmount)
            chkPaid.Checked = IIf(cParts.Paid, 1, 0)

            Notes_Text.Text = cParts.Notes
        Else  ' Can't find the record.  This is a problem.
            MessageBox.Show("Error locating parts order in database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            DisposeDA(cParts)
            Exit Sub
        End If

        DisposeDA(cParts)
    End Sub

    Private Sub SelectChargeBackOption(ByVal nVal As Integer)
        If nVal = 2 Then
            optCBCredit.Checked = True
        ElseIf nVal = 1 Then
            optCBDeduct.Checked = True
        Else ' 0
            optCBChargeBack.Checked = True
        End If
    End Sub

    Private Sub SelectStatus(ByVal Stat As String)
        Select Case UCase(Trim(Stat))
            Case "", "OPEN"  ' allow "" for clearing
                cboStatus.SelectedIndex = 0
            Case "CLOSED"
                cboStatus.SelectedIndex = 1
        End Select
    End Sub

    Public Function ClearServiceCall(Optional ByVal PreventEnableNavigation As Boolean = False) As Boolean
        ServiceCallNumber = 0

        lblServiceOrderNo.Text = ""
        lblFirstName.Text = ""
        lblLastName.Text = ""
        lblAddress.Text = ""
        lblAddress2.Text = ""
        lblCity.Text = ""
        lblZip.Text = ""
        lblTele.Text = ""
        lblTele2.Text = ""
        lblTele3.Text = ""
        UpdateTelephoneLabels("", "", "")

        ClearPartsOrder(PreventEnableNavigation)
        ClearServiceCall = True
    End Function

    Public Function ClearPartsOrder(Optional ByVal PreventEnableNavigation As Boolean = False) As Boolean
        On Error Resume Next
        PartsOrderID = 0
        MarginLine = 0

        lblPartsOrderNo.Text = ""
        lblMarginLine.Text = ""
        lblClaimDate.Text = ""
        SelectStatus("")

        SelectChargeBackOption(0)
        txtRepairCost.Text = FormatCurrency(0#)

        Notes_Text.Text = ""

        txtInvoiceNo.Text = ""
        dteClaimDate.Value = ""
        '  dteClaimDate.Value = date ' no longer clear to current date... BFH20050421
        txtSaleNo.Text = ""
        txtSaleNo.Tag = ""

        LoadVendor("")

        LoadStore(StoresSld)

        txtStyleNo.Text = ""
        txtDescription.Text = ""

        If Not PreventEnableNavigation Then EnableNavigation()
        ClearPartsOrder = True
    End Function

    Public Sub GetInvoiceInfoFromSaleNo(ByVal SaleNo As String)
        Dim SQL As String, RS As ADODB.Recordset
        SQL = "SELECT TOP 1 Misc as InvoiceNumber, DDate1 as InvoiceDate FROM Detail WHERE SaleNo = '" & SaleNo & "'"
        RS = GetRecordsetBySQL(SQL, , GetDatabaseInventory)
        On Error Resume Next
        If RS.EOF Then Exit Sub
        RS.MoveFirst()
        txtInvoiceNo.Text = RS("InvoiceNumber").Value
        dteClaimDate.Value = RS("InvoiceDate").Value

        DisposeDA(RS)
    End Sub

    ' LoadVendor(nVendorName) - send "" to clear..
    Private Function LoadVendor(ByVal nVendorName As String) As Boolean
        LoadVendorToServiceForm(Me, nVendorName)
    End Function

    Private Sub UpdateTelephoneLabels(ByVal Lbl1 As String, ByVal Lbl2 As String, ByVal Lbl3 As String)
        If Trim(Lbl1) = "" Then Lbl1 = "Tele: "
        If Trim(Lbl2) = "" Then Lbl2 = "Tele2: "
        If Trim(Lbl3) = "" Then Lbl3 = "Tele3: "
        If Microsoft.VisualBasic.Right(Trim(Lbl1), 1) <> ":" Then Lbl1 = Lbl1 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl2), 1) <> ":" Then Lbl2 = Lbl2 & ": "
        If Microsoft.VisualBasic.Right(Trim(Lbl3), 1) <> ":" Then Lbl3 = Lbl3 & ": "
        lblTele1Caption.Text = Lbl1
        lblTele2Caption.Text = Lbl2
        lblTele3Caption.Text = Lbl3
        Dim Longest As Integer
        Longest = Max(lblTele1Caption.Width, lblTele2Caption.Width, lblTele3Caption.Width)
        lblTele.Left = lblTele1Caption.Left + Longest + 60
        lblTele2.Left = lblTele1Caption.Left + Longest + 60
        lblTele3.Left = lblTele1Caption.Left + Longest + 60
    End Sub
End Class