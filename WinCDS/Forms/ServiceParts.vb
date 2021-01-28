Imports System.ComponentModel
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Public Class ServiceParts
    Private Mode As ServiceForMode
    Public CreateNewMode As ServiceForMode
    Dim Store As Integer
    'Dim ServiceCallNumber As Integer
    Public ServiceCallNumber As Integer
    Dim PartsOrderID As Integer
    Dim MarginLine As Integer
    Public Vendor As String
    Dim mCurrentCode As String
    Dim LocationVar As Integer
    Private OwnerVar As Form                         ' For showing as a child of Service, etc.
    Dim WithEvents mInvCkStyle As InvCkStyle
    Dim WithEvents mSelectStyle As InvCkStyle
    Dim mSelectedStyle As String
    Private Const Debugging As Boolean = False
    Private Const APText_Charged As String = "[""DEDUCT FROM INVOICE"" CHARGED TO AP]"
    'Private Const APText_Paid    As String = "[PAID IN AP]"
    Public ServicePartsFormLoaded As Boolean
    Dim FromSendChargeBackLetter As Boolean

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
                cmdMoveFirst.Width = 36
                'SetButtonImage(cmdMoveFirst, "previous")
                SetButtonImage(cmdMoveFirst, 7)
                cmdMoveFirst.Text = ""
                Navigate = True : Search = True
            Case 2  ' show 1
                cmdMoveFirst.Width = 195
                cmdMoveFirst.Text = "Browse Records"
                'cmdMoveFirst.Picture = Nothing
                cmdMoveFirst.Image = Nothing
                cmdMoveFirst.TextAlign = ContentAlignment.MiddleCenter
                Navigate = False : Search = False
            Case 3  ' show browse and search
                'cmdMoveFirst.Width = 80
                cmdMoveFirst.Width = 120
                cmdMoveFirst.Text = "Browse Records"
                'cmdMoveFirst.Picture = Nothing
                cmdMoveFirst.Image = Nothing
                cmdMoveFirst.TextAlign = ContentAlignment.MiddleCenter
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
            If cParts.InvoiceDate <> "" Then
                dteClaimDate.Value = cParts.InvoiceDate
            End If
            lblPartsOrderNo.Text = PO
            lblClaimDate.Text = Format(cParts.DateOfClaim, "MM/dd/yy")

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
        'dteClaimDate.Value = ""
        dteClaimDate.Value = Today
        'dteClaimDate.Value = date ' no longer clear to current date... BFH20050421
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
        lblTele.Left = lblTele1Caption.Left + Longest - 3
        lblTele2.Left = lblTele1Caption.Left + Longest - 3
        lblTele3.Left = lblTele1Caption.Left + Longest - 3
    End Sub

    Private Sub cmdEmail_Click(sender As Object, e As EventArgs) Handles cmdEmail.Click
        If IsDate(cmdEmail.Tag) Then
            If Math.Abs(DateDiff("s", cmdEmail.Tag, Now)) < 10 Then
                MessageBox.Show("Please wait a few moments for the email process to finish." & vbCrLf & " You will be notified of any success or failure.", "Please Wait!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        End If
        cmdEmail.Tag = Now
        frmEmail.EmailPartOrder(lblPartsOrderNo.Text, , txtVendorEmail.Text)
    End Sub

    Private Sub cmdPictures_Click(sender As Object, e As EventArgs) Handles cmdPictures.Click
        If Trim(lblPartsOrderNo.Text) = "" Then
            MessageBox.Show("Please select a valid Part Order.", "No Part Order", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        frmPictures.LoadPicturesByRef(frmPictures.dbPicType.dbpty_ServiceParts, lblPartsOrderNo.Text)
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        Dim P As Printer, N As String, NextY As Integer, Ty As Integer
        Dim A As Integer, B As Integer
        P = Printer

        P.FontName = "Arial"
        P.DrawWidth = 1
        P.FontSize = 10

        P.FontSize = 16 : PrintAligned("PARTS ORDER", VBRUN.AlignmentConstants.vbCenter, , 100, True)
        P.FontSize = 10

        PrintAutoMailingLetterHeader(txtStoreName.Text, txtStoreAddress.Text, txtStoreCity.Text, txtStorePhone.Text, txtVendorName.Text, txtVendorAddress.Text, txtVendorCity.Text, txtVendorTele.Text, True)

        P.Print("")
        P.Print("") : P.Print("") : P.Print("") : P.Print("")

        NextY = P.CurrentY

        A = 1000 : B = A + 2250
        Ty = NextY
        PrintAligned("Parts Order Number:", , A, Ty, True)
        PrintAligned(CStr(PartsOrderID), , B, Ty, True)
        Ty = P.CurrentY
        If Mode = ServiceForMode.ServiceMode_ForCustomer Then
            PrintAligned("Service Order Number:", , A, Ty, True)
            PrintAligned(CStr(ServiceCallNumber), , B, Ty, True)
            Ty = P.CurrentY
        End If
        PrintAligned("Date of Claim:", , A, Ty, True)
        PrintAligned(CStr(IfNullThenNilString(dteClaimDate.Value)), , B, Ty, True)
        Ty = P.CurrentY
        PrintAligned("Status:", , A, Ty, True)
        PrintAligned(CStr(cboStatus.Text), , B, Ty, True)
        Ty = P.CurrentY

        P.Print("")

        If Mode = ServiceForMode.ServiceMode_ForCustomer Then
            PrintAligned("Invoice No:", , A, Ty, True)
            PrintAligned(CStr(txtInvoiceNo.Text), , B, Ty, True)
            Ty = P.CurrentY
            PrintAligned("Invoice Date:", , A, Ty, True)
            PrintAligned(IfNullThenNilString(dteClaimDate.Value), , B, Ty, True)
            Ty = P.CurrentY
            PrintAligned("Sale Number:", , A, Ty, True)
            PrintAligned(txtSaleNo.Text, , B, Ty, True)
            Ty = P.CurrentY
        End If
        PrintAligned("Repair Cost:", , A, Ty, True)
        PrintAligned(txtRepairCost.Text, , B, Ty, True)
        Ty = P.CurrentY
        PrintAligned("Paid?", , A, Ty, True)
        PrintAligned(YesNo(chkPaid.Checked = True), , B, Ty, True)
        Ty = P.CurrentY
        N = DescribeChargeBackOption(GetChargeBackOption)
        PrintAligned("Reimbursement:", , A, Ty, True)
        PrintAligned(N, , B, Ty, True)

        P.Print("")
        NextY = P.CurrentY

        A = 1000 : B = 1080
        P.FontSize = 8 : PrintAligned("Style No", , A, NextY, True, True)
        P.FontSize = 10 : PrintAligned(txtStyleNo.Text, , B,, True)
        P.FontSize = 8 : PrintAligned("Description:", , 3000 + A, NextY, True, True)
        P.FontSize = 10 : PrintAligned(txtDescription.Text, , 3000 + B,, True)
        P.Print("")
        P.FontSize = 8 : PrintAligned("Notes:", , A, , True, True)
        P.FontSize = 10 : PrintAligned(WrapLongText(Notes_Text.Text, 100), , B,, True)

        P.Print("")
        NextY = P.CurrentY

        'imgPicture.Picture = FindDatabasePicture(0, cdspicType_PartsOrder, PartsOrderID)
        imgPicture.Image = FindDatabasePicture(0, cdspicType_PartsOrder, PartsOrderID)

        If Not imgPicture.Image Is Nothing Then
            'If imgPicture.Image <> 0 Then
            If imgPicture.Image IsNot Nothing Then
                MaintainPictureRatio(imgPicture, 10000, 5000, True)
                P.CurrentX = 0
                'P.PaintPicture(imgPicture.Image, (P.ScaleWidth - imgPicture.Width) / 2, P.CurrentY, imgPicture.Width, imgPicture.Height)
                P.PaintPicture(imgPicture.Image, (P.ScaleWidth - imgPicture.Width) / 2, P.CurrentY, 600, 400)
            End If
        End If

        P.EndDoc()
    End Sub

    Private Function GetChargeBackOption() As Integer
        Dim N As Integer
        If optCBChargeBack.Checked = True Then
            N = 0
        ElseIf optCBDeduct.Checked = True Then
            N = 1
        Else  ' optCBCredit
            N = 2
        End If
        GetChargeBackOption = N
    End Function

    Private Sub cmdPrintChargeBack_Click(sender As Object, e As EventArgs) Handles cmdPrintChargeBack.Click
        If GetPrice(txtRepairCost.Text) = 0 Then
            MessageBox.Show("Please enter a value into the Repair Cost field.", "Insufficient Data", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        FromSendChargeBackLetter = True
        'cmdSave.Value = True
        cmdSave_Click(cmdSave, New EventArgs)
        FromSendChargeBackLetter = False
        ServiceIntake.InitForm(ServiceCallNumber, PartsOrderID, GetChargeBackOption, txtRepairCost.Text, txtInvoiceNo.Text, Store, Vendor)
        'ServiceIntake.Show vbModal, Me
        ServiceIntake.ShowDialog(Me)
    End Sub

    Public Sub PrintedChargeBack(Optional ByVal Success As Boolean = False, Optional ByVal AsEmail As Boolean = False, Optional ByVal CBType As Integer = 0)
        If Success Then
            Notes_Text.Text = Notes_Text.Text & IIf(Len(Notes_Text.Text) > 0, vbCrLf, "") & "Charge Back Letter " & IIf(AsEmail, "Emailed", "Printed") & ": " & Today

            If CBType = 1 Then RecordRequestPayment()

            'cmdSave.Value = True
            cmdSave_Click(cmdSave, New EventArgs)
        End If
    End Sub

    Private Function APCharged() As Boolean
        APCharged = InStr(Notes_Text.Text, APText_Charged) > 0
    End Function

    Private Sub RecordRequestPayment()
        Dim A As Decimal
        Dim StoreNum As Integer, R As String, Mem As String
        Dim RET As Integer, RetMsg As String
        A = GetPrice(txtRepairCost.Text)
        StoreNum = IIf(StoreSettings.bPostToLoc1, 1, StoresSld)
        Mem = "Parts Order #" & lblPartsOrderNo.Text & IIf(Len(lblServiceOrderNo) > 0, " (SO#" & lblServiceOrderNo.Text & ")", "")

        If StoreSettings.bAPPost Then
            If Not APCharged() Then
                Notes_Text.Text = Notes_Text.Text & vbCrLf & APText_Charged & " " & Now
                'BFH20061127 - in as a negative amount because it is a credit
                SetAPTransaction(GetVendorCodeFromName(txtVendorName.Text), "PartOrd #" & lblPartsOrderNo.Text, Today, -A, Today, "11500", -A, , , -A, , , , , , , , , "JK")

                If UseQB(R) Then
                    QB_VendorQuery_Vendor(txtVendorName.Text, RET, RetMsg)
                    If RET <> 0 Then
                        MessageBox.Show("Vendor " & txtVendorName.Text & " does not exist in QuickBooks." & vbCrLf & "No invoice transaction will be posted.", "Vendor does not exist in QuickBooks", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        If Not QBCreateJournalEntry(Today, , , ,
              , , QueryGLQBAccountMap("10001"), A, Mem, , txtVendorName.Text, , QBLocationClassID(StoreNum, True),
              , , QueryGLQBAccountMap("11500"), A, Mem, , , , QBLocationClassID(StoreNum, True)) Then
                            MessageBox.Show("Repair Charge Posting to QB failed.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If
                Else
                    If QBWanted() Then MessageBox.Show("Could not record parts order charge to accounting QB." & vbCrLf & R, "QB Support Selected but not Available", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            End If
        End If
    End Sub

    Private Sub txtRepairCost_Enter(sender As Object, e As EventArgs) Handles txtRepairCost.Enter
        SelectContents(txtRepairCost)
    End Sub

    Private Sub txtRepairCost_Leave(sender As Object, e As EventArgs) Handles txtRepairCost.Leave
        On Error Resume Next
        txtRepairCost.Text = FormatCurrency(txtRepairCost.Text)
    End Sub

    Private Sub txtVendorName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles txtVendorName.SelectedIndexChanged
        LoadVendor(txtVendorName.Text)
    End Sub

    Private Sub txtVendorName_Validating(sender As Object, e As CancelEventArgs) Handles txtVendorName.Validating
        LoadVendor(txtVendorName.Text)
    End Sub

    Public Sub SetInvoiceInfoOnSaleNo(ByVal SaleNo As String, ByVal InvoiceNo As String, ByVal InvoiceDate As String)
        Dim SQL As String, IDC As String
        IDC = IIf(IsNothing(InvoiceDate), "NULL", "#" & InvoiceDate & "#")
        SQL = "UPDATE DETAIL SET Misc = '" & InvoiceNo & "', DDate1 = " & IDC & " WHERE SaleNo = '" & SaleNo & "'"
        ExecuteRecordsetBySQL(SQL, , GetDatabaseInventory)
    End Sub

    Private Sub dteClaimDate_CloseUp(sender As Object, e As EventArgs) Handles dteClaimDate.CloseUp
        If txtSaleNo.Tag <> "VALID" Then Exit Sub
        SetInvoiceInfoOnSaleNo(Val(txtSaleNo.Text), txtInvoiceNo.Text, dteClaimDate.Value)
    End Sub

    Private Sub txtInvoiceNo_Validating(sender As Object, e As CancelEventArgs) Handles txtInvoiceNo.Validating
        If txtSaleNo.Tag <> "VALID" Then Exit Sub
        SetInvoiceInfoOnSaleNo(Val(txtSaleNo.Text), txtInvoiceNo.Text, dteClaimDate.Value)
    End Sub

    Private Sub txtSaleNo_TextChanged(sender As Object, e As EventArgs) Handles txtSaleNo.TextChanged
        txtSaleNo.Tag = ""
    End Sub

    Private Sub cmdAddPart_Click(sender As Object, e As EventArgs) Handles cmdAddPart.Click
        ' Set mInvCkStyle's limit to items in Detail with MarginRn=0.
        If ServiceCallNumber > 0 Then
            ' Limit list to customer parts..
            Dim AddedML As Integer
            AddedML = AddOnAcc.GetMarginLine(ServiceCallNumber, Me)
            'Unload AddOnAcc
            AddOnAcc.Close()
            AddOnAcc = Nothing
            If AddedML <= 0 Then
                MessageBox.Show("Invalid item.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            LoadInfoFromMarginLine(AddedML)
        Else
            ' Limit list to parts in stock?
            ' Can't enforce that limit.  We aren't certain that everything is in Detail,
            ' but we'll try to get an invoice number if we can.
            mInvCkStyle = New InvCkStyle
            '    mInvCkStyle.ParentForm = Me.Name
            '    mInvCkStyle.LimitToStock = True   ' Allow any defined part.
            'mInvCkStyle.Show vbModal, Me 'ServiceParts
            mInvCkStyle.ShowDialog(Me)
            mInvCkStyle = Nothing
        End If
    End Sub

    Private Sub cmdMoveFirst_Click(sender As Object, e As EventArgs) Handles cmdMoveFirst.Click
        If Not LoadRelativePartsOrder(-1, True, Mode = ServiceForMode.ServiceMode_ForCustomer) And cmdMoveFirst.Text <> "<<" Then
            MessageBox.Show("There were no other parts orders for this Service Call.", "Not Found!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        EnableNavigation()
    End Sub

    Private Sub cmdMoveLast_Click(sender As Object, e As EventArgs) Handles cmdMoveLast.Click
        LoadRelativePartsOrder(1, True, Mode = ServiceForMode.ServiceMode_ForCustomer)
    End Sub

    Private Sub cmdMoveNext_Click(sender As Object, e As EventArgs) Handles cmdMoveNext.Click
        LoadRelativePartsOrder(1, , Mode = ServiceForMode.ServiceMode_ForCustomer)
    End Sub

    Private Sub cmdMovePrevious_Click(sender As Object, e As EventArgs) Handles cmdMovePrevious.Click
        LoadRelativePartsOrder(-1, , Mode = ServiceForMode.ServiceMode_ForCustomer)
    End Sub

    Private Sub cmdNext_Click(sender As Object, e As EventArgs) Handles cmdNext.Click
        SelectMode(CreateNewMode, True)
        If CreateNewMode = ServiceForMode.ServiceMode_ForCustomer Then
            ClearPartsOrder()
        Else
            ClearServiceCall()
        End If
    End Sub

    Private Sub cmdMoveSearch_Click(sender As Object, e As EventArgs) Handles cmdMoveSearch.Click
        Dim X As Integer
        X = Val(InputBox("Search for Order Number:", "New Service Parts Order Number"))
        If X <= 0 Then Exit Sub
        LoadPartsOrder(X)
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Dim cParts As clsServicePartsOrder
        cParts = New clsServicePartsOrder
        ' This could be a new parts order or an update.
        ' If it's an update, load the existing parts Order.
        If PartsOrderID <> 0 Then
            If cParts.Load(CStr(PartsOrderID), "#ServicePartsOrderNo") Then
                ' We've got the record, it's okay.
            Else
                ' Can't find the record.  This is a problem.
                MessageBox.Show("Error locating parts order in database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        End If

        ' Then regardless, set the form variables into it.
        cParts.MarginLine = MarginLine
        cParts.Style = txtStyleNo.Text
        cParts.Desc = txtDescription.Text
        cParts.Store = Store

        cParts.Vendor = txtVendorName.Text
        '    If txtVendorName.ListIndex >= 0 Then
        '      cParts.VendorNo = txtVendorName.ItemData(txtVendorName.ListIndex)
        '    Else
        '      cParts.VendorNo = ""
        '    End If
        cParts.VendorAddress = txtVendorAddress.Text
        cParts.VendorCity = txtVendorCity.Text
        cParts.VendorTele = txtVendorTele.Text

        cParts.ServiceOrderNo = ServiceCallNumber
        cParts.DateOfClaim = Today
        cParts.Status = cboStatus.Text
        cParts.Notes = Notes_Text.Text

        cParts.ChargeBackType = GetChargeBackOption()
        cParts.ChargeBackAmount = GetPrice(txtRepairCost.Text)
        cParts.Paid = chkPaid.Checked = True

        If MarginLine <> 0 And ServiceCallNumber <> 0 Then
            If cParts.NoteID = 0 Then
                cParts.NoteID = CreateServiceNote(MarginLine, ServiceCallNumber, "Ordered Part, Status=" & cboStatus.Text, , 1)
            Else
                SetServiceNoteText(cParts.NoteID, "Ordered Part, Status=" & cboStatus.Text)
            End If
        Else
            cParts.NoteID = 0
        End If

        cParts.InvoiceDate = IfNullThenNilString(dteClaimDate.Value)
        cParts.InvoiceNo = txtInvoiceNo.Text

        ' Save, and grab the Autonumber (in case it's a new record).
        cParts.Save()
        PartsOrderID = cParts.ServicePartsOrderNo

        'Dim rsMax As New ADODB.Recordset
        'rsMax = GetRecordsetBySQL("Select max(ServicePartsOrderNo) from ServicePartsOrder", True, GetDatabaseAtLocation)
        'If Not rsMax.EOF And Not rsMax.BOF Then
        '    PartsOrderID = rsMax(0).Value
        '    cParts.ServicePartsOrderNo = PartsOrderID
        'End If

        lblPartsOrderNo.Text = cParts.ServicePartsOrderNo
        lblClaimDate.Text = Format(cParts.DateOfClaim, "MM/dd/yy")

        ' Also save the actual parts.. We need the autonumber first!

        EnableNavigation()

        DisposeDA(cParts)
    End Sub

    Public Sub InitStatusList()
        cboStatus.Items.Clear()
        cboStatus.Items.Insert(0, "Open")
        cboStatus.Items.Insert(1, "Closed")
        cboStatus.SelectedIndex = 0
    End Sub

    Private Sub ServiceParts_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ServicePartsFormLoaded = True Then Exit Sub

        SetButtonImage(cmdSave, 2)
        SetButtonImage(cmdPrint, 19)
        SetButtonImage(cmdEmail, 1)
        SetButtonImage(cmdNext, 6)
        SetButtonImage(cmdMenu, 9)
        SetButtonImage(cmdPictures, 26)
        SetButtonImage(cmdMoveFirst, 7)
        SetButtonImage(cmdMovePrevious, 4)
        SetButtonImage(cmdMoveNext, 5)
        SetButtonImage(cmdMoveLast, 6)

        ' general initialization
        ' these 2 lables are just for displaying debugging information
        '  lblMarginLinelbl.Visible = False
        '  lblMarginLine.Visible = False

        CreateNewMode = ServiceForMode.ServiceMode_ForCustomer
        InitStatusList()
        LoadMfgNamesIntoComboBox(txtVendorName, , True, True)
        LoadStoresIntoComboBox(cboStores, , True)
        SelectMode(, True, True) 'default to order for customer

        ClearServiceCall()

        If Debugging Then
            lblMarginLine.Visible = Debugging
            lblMarginLinelbl.Visible = Debugging
        End If
    End Sub

    Private Sub ServiceParts_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'If UnloadMode = vbFormControlMenu Then
        '    cmdMenu.Value = True
        'End If

    End Sub

    Private Sub mInvCkStyle_CancelClicked(ByRef Override As Boolean) Handles mInvCkStyle.CancelClicked
        Override = True   ' Don't InvCkStyle continue processing.
        'Unload mInvCkStyle
        mInvCkStyle.Close()
    End Sub

    Private Sub mInvCkStyle_OKClicked(ByRef Override As Boolean, Picked As String, IsNew As Boolean) Handles mInvCkStyle.OKClicked
        Override = True  ' No matter what, don't InvCkStyle continue processing.
        If IsNew Then Exit Sub     ' We only allow existing styles.
        ' Bring up AddOnAcc with a list of items, and go with the result.
        Dim AddedDL As Integer
        AddedDL = AddOnAcc.GetDetailLine(Picked, Me)
        If AddedDL <= 0 Then
            MessageBox.Show("Invalid item.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Override = True
            Exit Sub
        End If

        txtStyleNo.Text = Picked  ' This is temporary!  We will have to get Detail for the item.
        Notes_Text.Text = "Detail Line " & AddedDL
        'Unload mInvCkStyle
        mInvCkStyle.Close()
    End Sub

    Private Sub cmdMenu_Click(sender As Object, e As EventArgs) Handles cmdMenu.Click
        If Owner Is Nothing Then
            modProgramState.Order = ""
            MainMenu.Show()
        Else
            Owner.Show()
            If Owner.Name = Service.Name Then Service.PartsOrderFormClosed()
            Owner = Nothing
        End If
        'Unload Me
        Me.Close()
    End Sub

    Private Sub cboStores_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboStores.SelectedIndexChanged
        'LoadStore(cboStores.itemData(cboStores.ListIndex))
        LoadStore(CType(cboStores.Items(cboStores.SelectedIndex), ItemDataClass).ItemData)
    End Sub

    Private Sub mSelectStyle_CancelClicked(ByRef Override As Boolean) Handles mSelectStyle.CancelClicked
        Override = True
        mSelectedStyle = ""
        'Unload mSelectStyle
        mSelectStyle.Close()
    End Sub

    Private Sub mSelectStyle_OKClicked(ByRef Override As Boolean, Picked As String, IsNew As Boolean) Handles mSelectStyle.OKClicked
        Override = True
        mSelectedStyle = Picked
        'Unload mSelectStyle
        mSelectStyle.Close()
    End Sub

    Private Sub SelectStyle(ByRef Style As String, ByRef Description As String, ByRef Vendor As String)
        mSelectStyle = New InvCkStyle

        'mSelectStyle.Show vbModal, Me
        mSelectStyle.ShowDialog(Me)
        If Len(mSelectedStyle) > 0 Then
            Style = mSelectedStyle
            GetInfoFromStyle(Style, Description, Vendor, "")
        End If
    End Sub

    Private Sub txtStyleNo_Enter(sender As Object, e As EventArgs) Handles txtStyleNo.Enter
        Dim ST As String, Des As String, Ven As String
        If Mode <> ServiceForMode.ServiceMode_ForCustomer Then
            Notes_Text.Select()  ' this, among other things, prevents bad loops
            SelectStyle(ST, Des, Ven)
            If Len(ST) > 0 Then
                txtStyleNo.Text = ST
                txtDescription.Text = Des
                LoadVendor(Ven)
            End If
        End If
    End Sub

    Public Function ChargeBackTypeDesc(ByVal Typ As Integer, Optional ByVal LDesc As Boolean = False) As String
        Select Case Typ
            Case 0 : ChargeBackTypeDesc = IIf(LDesc, "Charging Back", "Charge Back")
            Case 1 : ChargeBackTypeDesc = IIf(LDesc, "Deducting", "Deduct")
            Case 2 : ChargeBackTypeDesc = IIf(LDesc, "Requesting a Credit", "Credit")
            Case Else : ChargeBackTypeDesc = "???"
        End Select
    End Function
End Class