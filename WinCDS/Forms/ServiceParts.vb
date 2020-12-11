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

    Public Sub LoadInfoFromMarginLine(Optional ByVal ML As Long = -1, Optional ByVal HideSaleNo As Boolean = False)
        Dim Margin As CGrossMargin
  
  Set Margin = New CGrossMargin
  
  If ML <> -1 Then MarginLine = ML

        If Margin.Load(CStr(MarginLine), "#MarginLine") Then
            txtStyleNo = Margin.Style
            txtDescription = Margin.Desc
            txtSaleNo = Margin.SaleNo
            GetInvoiceInfoFromSaleNo Margin.SaleNo
    txtSaleNo.Tag = IIf(txtSaleNo = "", "", "VALID")
            LoadVendor Margin.Vendor
'    Notes_Text.Text = "Margin Line " & AddedML  ' removed:  bfh20050218
        Else
            MsgBox "Error: Can't load GrossMargin item.", vbCritical, "Error"
    MarginLine = 0
        End If

        lblMarginLine = MarginLine

        If MarginLine = 0 Then
            txtStyleNo = ""
            txtDescription = ""
            txtSaleNo = ""
        End If
        If HideSaleNo Then
            txtSaleNo.Visible = False
            lblSaleNo.Visible = False
        End If

        DisposeDA Margin
End Sub

    Public Function LoadServiceCall(ByVal sC As Long) As Boolean
        ' Load as much as we can from Service Call Number (vendor, cust, etc).
        ServiceCallNumber = sC
        lblServiceOrderNo.Caption = sC
        Dim nSC As clsServiceOrder
        Dim MailRec As clsMailRec
  
  Set nSC = New clsServiceOrder
  If nSC.Load(CStr(sC), "#ServiceOrderNo") Then
            ' Fill in necessary SC-related information.
            ' Customer name and address
            LoadServiceCall = True
    ' Look up mail record for this information:
    Set MailRec = New clsMailRec
    If MailRec.Load(nSC.MailIndex, "#Index") Then
                lblFirstName.Caption = MailRec.First
                lblLastName.Caption = MailRec.Last
                lblAddress.Caption = MailRec.Address
                lblAddress2.Caption = MailRec.AddAddress
                lblCity.Caption = MailRec.City
                lblZip.Caption = MailRec.Zip
                lblTele.Caption = DressAni(CleanAni(MailRec.Tele))
                lblTele2.Caption = DressAni(CleanAni(MailRec.Tele2))
                Dim Mail2 As MailNew2
                modMail.Mail2_GetAtIndex MailRec.Index, Mail2
      lblTele3.Caption = Mail2.Tele3
                UpdateTelephoneLabels MailRec.PhoneLabel1, MailRec.PhoneLabel2, Mail2.PhoneLabel3
    Else ' Can't load mail record.
                lblLastName.Caption = "Missing Customer Information"
                UpdateTelephoneLabels "", "", ""
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

        DisposeDA nSC, MailRec
End Function

    Public Function LoadRelativePartsOrder(ByVal Dir As Long, Optional ByVal Max As Boolean = False, Optional ByVal RestrictToCurrentServiceCall As Boolean = True) As Boolean
        Dim SQL As String, BaseRestrict As String, DirS As String, DirP As String
        Dim RS As Recordset, NewID As Long

        If Dir = 0 Then Exit Function
        BaseRestrict = "WHERE (TRUE=TRUE)" ' allows adding additional " AND ..." clauses w/o checks
        If CreateNewMode = ServiceMode_ForCustomer And ServiceCallNumber <> 0 And RestrictToCurrentServiceCall Then
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
  Set RS = GetRecordsetBySQL(SQL, , GetDatabaseAtLocation())
  
  On Error GoTo NoID
        NewID = 0
        NewID = RS("ServicePartsOrderNo")
        If NewID <> 0 Then
            ClearServiceCall True
    LoadPartsOrder NewID
  End If
        LoadRelativePartsOrder = True
        DisposeDA RS
NoID:
    End Function

End Class